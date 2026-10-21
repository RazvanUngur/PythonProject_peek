# =============================================================================
# parallel_utils.py — Infrastructură comună pentru procesarea pe mai multe nuclee
# =============================================================================
# Folosit de bin_parser.py și log_parser.py.
#
# Principii:
#   * Procese (nu fire): pandas/openpyxl sunt cod Python pur, blocat de GIL.
#   * Contextul "spawn" peste tot (identic pe Windows/Linux; sigur față de un
#     proces principal cu interfață grafică și fire multiple).
#   * Procesele worker NU ating niciodată SQLite. Toate citirile/scrierile din
#     DB se fac în procesul principal (baza e pe rețea; un singur scriitor).
#   * Un worker nu ridică excepții către părinte: întoarce un dict
#     {"ok", "result"|"error", "log"}; print()-urile lui sunt captate în "log"
#     și afișate de procesul principal.
#   * Dacă un worker cade (ex. memorie insuficientă), contorul se reia automat
#     în procesul principal.
#
# Numărul de procese:
#   implicit  = min(nuclee - 1, 6)   (pe 8 nuclee → 6)
#   suprascriere:  config.MAX_WORKERS = N   sau variabila de mediu CESTRIN_WORKERS=N
#   1 = fără paralelism (comportament identic cu versiunea veche).
# =============================================================================

import os
import io
import contextlib
import multiprocessing
import time
import traceback
from contextlib import contextmanager
from concurrent.futures import ProcessPoolExecutor, Future, wait, FIRST_COMPLETED

try:
    import config as _cfg
except Exception:            # pragma: no cover
    _cfg = None


class StopRequested(Exception):
    """Utilizatorul a apăsat Anulează."""


# ─────────────────────────────────────────────────────────────────────────────
# Număr de procese
# ─────────────────────────────────────────────────────────────────────────────
def get_worker_count(n_tasks):
    forced = os.environ.get("CESTRIN_WORKERS") or getattr(_cfg, "MAX_WORKERS", None)
    try:
        forced = int(forced) if forced else None
    except (TypeError, ValueError):
        forced = None
    if forced is not None:
        n = max(1, forced)
    else:
        n = max(1, min((os.cpu_count() or 2) - 1, 6))
    return max(1, min(n, int(n_tasks)))


# ─────────────────────────────────────────────────────────────────────────────
# Executor
# ─────────────────────────────────────────────────────────────────────────────
class _SerialExecutor:
    """Rulează sarcinile imediat, în procesul curent (fallback / 1 worker)."""
    def submit(self, fn, *args, **kwargs):
        f = Future()
        try:
            f.set_result(fn(*args, **kwargs))
        except Exception as e:
            f.set_exception(e)
        return f

    def shutdown(self, wait=True, cancel_futures=False):
        pass


class PoolHandle:
    def __init__(self, n_workers):
        self.n_workers = max(1, int(n_workers))
        if self.n_workers > 1:
            ctx = multiprocessing.get_context("spawn")
            self._ex = ProcessPoolExecutor(max_workers=self.n_workers, mp_context=ctx)
            print(f"[Paralelizare] pornesc {self.n_workers} procese "
                  f"(nuclee detectate: {os.cpu_count()})")
        else:
            self._ex = _SerialExecutor()
            print(f"[Paralelizare] 1 proces (serial) — "
                  f"nuclee detectate: {os.cpu_count()}")

    @property
    def is_parallel(self):
        return self.n_workers > 1

    def submit(self, fn, *args, **kwargs):
        try:
            return self._ex.submit(fn, *args, **kwargs)
        except Exception as e:      # ex. BrokenProcessPool: rulăm local
            print(f"[Paralelizare][ATENȚIE] trimiterea către un proces worker "
                  f"a eșuat imediat ({e!r}); rulez acest job în procesul "
                  f"principal.")
            f = Future()
            try:
                f.set_result(fn(*args, **kwargs))
            except Exception as e2:
                f.set_exception(e2)
            return f

    def close(self):
        try:
            self._ex.shutdown(wait=True)
        except Exception:
            pass

    def abort(self):
        """Oprire imediată: anulează ce nu a început și oprește procesele."""
        procs = []
        try:
            procs = list((getattr(self._ex, "_processes", None) or {}).values())
        except Exception:
            pass
        try:
            self._ex.shutdown(wait=False, cancel_futures=True)
        except TypeError:
            self._ex.shutdown(wait=False)
        except Exception:
            pass
        for p in procs:
            try:
                p.terminate()
            except Exception:
                pass


# ─────────────────────────────────────────────────────────────────────────────
# Apel "sigur" — rulează în worker, nu ridică excepții, captează print()
# ─────────────────────────────────────────────────────────────────────────────
def safe_call(fn, *args):
    in_child = multiprocessing.current_process().name != "MainProcess"
    buf = io.StringIO()
    try:
        if in_child:
            with contextlib.redirect_stdout(buf), contextlib.redirect_stderr(buf):
                res = fn(*args)
        else:
            res = fn(*args)
        return {"ok": True, "result": res, "log": buf.getvalue()}
    except Exception:
        return {"ok": False, "error": traceback.format_exc(), "log": buf.getvalue()}


def submit_safe(pool, fn, *args):
    return pool.submit(safe_call, fn, *args)


def get_safe(fut, fn, *args):
    """Rezultatul unui future creat cu submit_safe. Dacă procesul worker a
    căzut (BrokenProcessPool etc.), reia sarcina în procesul principal —
    dar afișează motivul, ca să nu treacă neobservat (altfel rezultatul
    iese corect și pare că totul e în regulă, doar că fără niciun câștig
    de viteză)."""
    try:
        return fut.result()
    except Exception:
        print("[Paralelizare][ATENȚIE] un job a eșuat în procesul worker și "
              "a fost refăcut în procesul principal (fără paralelism pentru "
              "acest job). Motiv:")
        traceback.print_exc()
        return safe_call(fn, *args)


def wait_one(fut, stop_event=None, poll=0.25):
    """Așteaptă un future, verificând periodic Anulează."""
    while True:
        if stop_event is not None and stop_event.is_set():
            raise StopRequested()
        done, _ = wait([fut], timeout=poll)
        if done:
            return


# ─────────────────────────────────────────────────────────────────────────────
# Orchestrator pentru generarea rapoartelor Excel
# ─────────────────────────────────────────────────────────────────────────────
class ReportRunner:
    """
    Trimite joburile Excel (unul per contor) către procese și, pe măsură ce se
    termină, în procesul principal:
      * aplică scrierile în trafic_mzl / trafic_mza,
      * adaugă timpul instrumentat pentru foaia "Date detaliate prelucrate",
      * apelează progress_callback(site_id, n_ore, idx, total).
    """
    def __init__(self, pool, stop_event, progress_callback, total, warn=None):
        self.pool = pool
        self.stop_event = stop_event
        self.progress_callback = progress_callback
        self.total = total
        self.warn = warn or print
        self._pending = {}       # future -> (job, meta)
        self._results = []       # (order, rezultat)
        self._tmp_files = []
        self._n_done = 0
        self.errors = []

    # -- trimitere ----------------------------------------------------------
    def submit(self, job, order, rezultat, site_id, randuri):
        from excel_report import build_counter_report_task
        self._tmp_files.append(job["tmp_fn"])
        fut = submit_safe(self.pool, build_counter_report_task, job)
        self._pending[fut] = (job, {"order": order, "rezultat": rezultat,
                                    "site_id": site_id, "randuri": randuri})
        self.poll()

    # -- colectare ----------------------------------------------------------
    def poll(self):
        for f in [f for f in self._pending if f.done()]:
            self._handle(f)

    def wait_all(self):
        """False dacă utilizatorul a anulat."""
        while self._pending:
            if self.stop_event is not None and self.stop_event.is_set():
                return False
            wait(list(self._pending), timeout=0.25, return_when=FIRST_COMPLETED)
            self.poll()
        return True

    def _handle(self, fut):
        from excel_report import build_counter_report_task, apply_db_writes, add_ddp_perf
        job, meta = self._pending.pop(fut)
        site_id = meta["site_id"]

        res = get_safe(fut, build_counter_report_task, job)
        if res.get("log"):
            print(res["log"], end="" if res["log"].endswith("\n") else "\n")

        if not res["ok"]:
            self.errors.append(site_id)
            last = (res.get("error") or "").strip().splitlines()
            self.warn(f"  ✗ [{site_id}] eroare la generarea raportului Excel: "
                      f"{last[-1] if last else 'necunoscută'}")
            print(res.get("error", ""))
            return

        out = res["result"]
        try:
            apply_db_writes(out.get("db_writes", []))
        except Exception as e:
            print(f"[WARN] Scriere trafic_mzl/mza [{site_id}]: {e}")
        add_ddp_perf(out.get("ddp_seconds", 0.0))

        self._n_done += 1
        self._results.append((meta["order"], meta["rezultat"]))
        if self.progress_callback:
            try:
                self.progress_callback(site_id, meta["randuri"], self._n_done, self.total)
            except Exception:
                pass

    # -- final --------------------------------------------------------------
    def finish(self):
        """Lista rezultatelor în ordinea inițială, sau None dacă s-a anulat."""
        if not self.wait_all():
            self.abort()
            return None
        self.pool.close()
        return [r for _, r in sorted(self._results, key=lambda x: x[0])]

    def abort(self):
        self.pool.abort()
        self._pending.clear()
        for p in self._tmp_files:          # fișiere temporare rămase
            try:
                if os.path.exists(p):
                    os.remove(p)
            except Exception:
                pass

# ─────────────────────────────────────────────────────────────────────────────
# Instrumentare separată pentru timpul petrecut în SQLite (citiri + scrieri),
# pe tip de operație, ca să se vadă exact ce anume domină în faza
# "Procesare + SQLite + Excel": scrierea (upsert), citirea istoricului
# (get_hourly_df), sau cererile mici (MZL manual, surse fișiere).
# ─────────────────────────────────────────────────────────────────────────────
_DB_PERF = {}   # {etichetă: secunde_cumulate}

# Ordinea în care apar etichetele în raport (nu alfabetic)
DB_PERF_LABELS = [
    ("scriere (upsert_hourly_df)",        "scriere"),
    ("citire (get_hourly_df)",            "citire"),
    ("MZL manual (get_mzl_manual)",       "mzl_manual"),
    ("surse fișiere (get_source_files*)", "surse"),
]


def reset_db_perf():
    global _DB_PERF
    _DB_PERF = {}


def add_db_perf(label, seconds):
    _DB_PERF[label] = _DB_PERF.get(label, 0.0) + float(seconds or 0.0)


def get_db_perf_seconds():
    """Total, pentru compatibilitate cu raportul existent."""
    return sum(_DB_PERF.values())


def get_db_perf_breakdown():
    """[(etichetă_afișată, secunde), ...] în ordinea DB_PERF_LABELS, doar
    etichetele cu timp > 0."""
    out = []
    for label_afisat, key in DB_PERF_LABELS:
        secs = _DB_PERF.get(key, 0.0)
        if secs > 0:
            out.append((label_afisat, secs))
    return out


@contextmanager
def timed_db(label):
    """Folosire: with timed_db("scriere"): tdb.upsert_hourly_df(...)
    Etichete valide: scriere, citire, mzl_manual, surse."""
    t0 = time.perf_counter()
    try:
        yield
    finally:
        add_db_perf(label, time.perf_counter() - t0)

