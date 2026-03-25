# =============================================================================
# bin_parser.py — Parsare fișiere .bin PEEK/Sabre
# =============================================================================
# Exportă:
#   process_peek_bin(filepath)   → (df, site_id, start_date, n_lanes)
#   quick_scan_bin(filepath)     → (site_id, n_records, n_lanes)
#   process_multiple_files(...)  → [{ path, id, randuri, b1, b2, n_lanes }, ...]
# =============================================================================

import os
import re
import struct
import pandas as pd
from datetime import datetime, timedelta

from config import (
    CENTRAL_FILE_NAME, CENTRAL_FILE_FOLDER,
    MIN_ORE_ZI, BAND_COLORS,
)
from excel_report import add_charts_and_formatting

def process_peek_bin(filepath):
    """
    Parsează un fișier .bin PEEK/Sabre și returnează (df, site_id, start_date, n_lanes).

    Detectare număr benzi: byte11 din header RO23 / 2
        4 → 2 benzi | 8 → 4 benzi | 12 → 6 benzi

    Structura record decodat:

      2 benzi (rec_size=34, ora=[0]):
          [3:18] = B1_cls1..15,  [18:33] = B2_cls1..15

      4 benzi – VARIANTA A (fără prefix totaluri):
          [8:23]=B1, [23:38]=B2, [38:53]=B3, [53:66]=B4 (cls1..13 + 2 zerouri)
          Detectare: mediana(vals[10]) == 0

      4 benzi – VARIANTA B (cu prefix totaluri la [8:12]):
          [12:27]=B1, [27:42]=B2, [42:57]=B3, [57:66]=B4_cls1..9 curent
          + B4_cls10..15 din recordul următor (cross-record)
          Detectare: mediana(vals[10]) > 10

      6 benzi (rec_size=98, ora=[11]):
          [12:27]=B1, [27:42]=B2, [42:57]=B3, [57:72]=B4, [72:87]=B5, [87:98]=B6(11cls)
    """
    base_name        = os.path.basename(filepath)
    name_without_ext = os.path.splitext(base_name)[0]
    is_sabre         = not name_without_ext[-1].isdigit()

    with open(filepath, "rb") as f:
        raw = f.read()

    # ── SITE ID — metodă inteligentă din header binar ────────────────────────
    try:
        header_text = "".join([chr(b) if 32 <= b <= 126 else " " for b in raw[0:150]])
        matches = re.findall(r'000+(\d{4})', header_text)
        if matches:
            site_id = matches[-1]
        else:
            # fallback: din numele fișierului
            m = re.search(r'(\d{4})', base_name)
            site_id = m.group(1) if m else "0"
    except Exception:
        site_id = "0"

    # ── Data de start din header ─────────────────────────────────────────────
    day_start    = raw[0x05]
    month_start  = raw[0x06]
    year_start   = raw[0x07] + 2000
    hour_start   = raw[0x04] if raw[0x04] <= 23 else 0   # ora de start (ignorată anterior)
    current_date = datetime(year_start, month_start, day_start, hour_start, 0)

    # ── Detectare număr benzi din header RO23 ────────────────────────────────
    pos_ro23 = raw.find(b'RO23')
    if pos_ro23 >= 0:
        raw_lanes = raw[pos_ro23 + 11] // 2      # 4→2, 8→4, 12→6
        n_lanes   = raw_lanes if raw_lanes in (2, 4, 6) else 2
    else:
        n_lanes = 2
    start_pos = (pos_ro23 + 68) if pos_ro23 >= 0 else 68

    # ── Format Ro04R (simplificat, mereu 2 benzi) ────────────────────────────
    is_ro04r = (raw.find(b'Ro04R') != -1) and (pos_ro23 == -1)

    if is_ro04r:
        # Decodare specifică Ro04R: totaluri simple per bandă
        def _decode_val_ro04r(data, i):
            b = data[i]
            if b >= 128:
                if i + 1 < len(data):
                    return (b & 0x7F) * 256 + data[i + 1], i + 2
                return None, i + 1
            return b, i + 1

        ro04r_idx = raw.find(b'Ro04R')
        pos = ro04r_idx + 14
        rows = []
        while pos < len(raw):
            b = raw[pos]
            if b == 0x00:
                pos += 1; continue
            if b > 23:
                pos += 1; continue
            hour_val = b
            pos += 1
            vals = []
            i = pos
            for _ in range(4):
                if i >= len(raw): break
                v, i = _decode_val_ro04r(raw, i)
                if v is not None: vals.append(v)
            if len(vals) < 4: break
            pos = i
            tot_b1, tot_b2 = vals[0], vals[1]
            c_b1 = [0] * 15; c_b1[14] = tot_b1
            c_b2 = [0] * 15; c_b2[14] = tot_b2

            if not rows:
                timestamp = current_date.replace(hour=hour_val, minute=0) - timedelta(hours=1)
            else:
                timestamp = rows[-1]["Timestamp"] + timedelta(hours=1)

            row = {"Contor": site_id, "Timestamp": timestamp,
                   "Data_Ora": timestamp.strftime("%d.%m.%Y %H:%M"), "N_Benzi": 2}
            for idx, val in enumerate(c_b1, 1): row[f"B1_Clasa_{idx}"] = val
            row["Total_B1"] = tot_b1
            for idx, val in enumerate(c_b2, 1): row[f"B2_Clasa_{idx}"] = val
            row["Total_B2"] = tot_b2
            # benzi 3+4 = zero pentru compatibilitate (nu există în Ro04R)
            for bn in range(3, 5):
                for idx in range(1, 16): row[f"B{bn}_Clasa_{idx}"] = 0
                row[f"Total_B{bn}"] = 0
            row["Total_General"] = tot_b1 + tot_b2
            rows.append(row)

        df = pd.DataFrame(rows)
        n_lanes = 2

    else:
        # ── Decodare completă (toate valorile) ───────────────────────────────
        all_vals = []
        i = start_pos
        while i < len(raw):
            b = raw[i]
            if b >= 128:
                if i + 1 < len(raw):
                    all_vals.append((b & 0x7F) * 256 + raw[i + 1])
                    i += 2
                else:
                    break
            else:
                all_vals.append(b)
                i += 1

        # ── Parametri structură per număr benzi ──────────────────────────────
        if n_lanes == 2:
            rec_size          = 34
            hour_idx          = 0
            lane_start        = 3
            has_totals_prefix = False
        elif n_lanes == 4:
            rec_size  = 66
            hour_idx  = 7
            # Detectare variantă A vs B: mediana vals[10] pe primele 24 recorduri
            sample_v10 = sorted([all_vals[s * 66 + 10]
                                  for s in range(min(24, len(all_vals) // 66))
                                  if s * 66 + 10 < len(all_vals)])
            median_v10        = sample_v10[len(sample_v10) // 2] if sample_v10 else 0
            has_totals_prefix = (median_v10 > 10)
            lane_start        = 12 if has_totals_prefix else 8
        else:  # 6 benzi — format ADR 3000 Clasificator cu varint encoding
            # Fișierele 6 benzi folosesc un format diferit:
            # • Valori >= 128 sunt codificate pe 2 octeți: (byte-128)*256 + next_byte
            # • Fiecare record = 2 octeți header + 6 benzi × 15 clase (varint)
            # • Dimensiunea recordului este variabilă (tipic 90-95 octeți)
            # • Primul record decodat este un bloc index, se sare
            # • Timestamp pornește din current_date (inclusiv ora din raw[4])
            rec_size          = 98   # valoare fallback, nu se folosește în ramura 6-benzi
            hour_idx          = 11
            lane_start        = 16
            has_totals_prefix = False
            # >>> Parsare specială 6-benzi se face MAI JOS, după blocul is_sabre <<<

        # ── Detecție Sabre din conținut ──────────────────────────────────────
        def _detect_sabre(all_vals, rec_size):
            markers, step = [], 0
            while step * rec_size + 5 <= len(all_vals) and len(markers) < 48:
                base = step * rec_size
                if all_vals[base] > 23:
                    step += 1; continue
                if base + 4 < len(all_vals):
                    markers.append(all_vals[base + 4])
                step += 1
            if len(markers) < 24 or any(v > 23 for v in markers):
                return False
            return sum(1 for k in range(len(markers) - 1)
                       if markers[k + 1] == (markers[k] + 1) % 24) >= 20

        def _detect_totalizator(raw, start_pos):
            """Detectează fișiere totalizator: hour + B1 + B2 + B1dup + B2dup + 0x00 per record.
            Condiție suplimentară: cel puțin un record trebuie să aibă valori non-zero.
            Fișierele clasificator offline au toate recordurile zero și trec
            accidental testul structural (fiecare byte 0 este varint valid + terminator).
            """
            pos = start_pos
            matches = 0
            any_nonzero = False
            for _ in range(min(20, (len(raw) - start_pos) // 6)):
                if pos >= len(raw): break
                h = raw[pos]
                if h > 23: return False
                pos += 1
                vals = []
                for _ in range(4):  # 4 valori varint
                    if pos >= len(raw): return False
                    b = raw[pos]
                    if b >= 128:
                        if pos + 1 >= len(raw): return False
                        vals.append((b & 0x7F) * 256 + raw[pos + 1])
                        pos += 2
                    else:
                        vals.append(b)
                        pos += 1
                if pos >= len(raw) or raw[pos] != 0: return False
                pos += 1  # skip terminator 0x00
                matches += 1
                if any(v > 0 for v in vals):
                    any_nonzero = True
                    # Totalizatorul stochează fiecare valoare de două ori: [B1, B2, B1dup, B2dup]
                    # Un clasificator care trece accidental testul structural NU are duplicate.
                    if vals[0] != vals[2] or vals[1] != vals[3]:
                        return False  # nu e totalizator real
            # Un fișier clasificator offline (trafic zero) trece testul structural
            # dar NU are niciun record non-zero → nu e totalizator real.
            return matches >= 5 and any_nonzero

        # ── Detecție variant2b: rec_size=34 în all_vals, hour@4, B2_total@6 ──
        # Format: rec[0..3]=header(0), rec[4]=hour, rec[6]=B2_total,
        #         rec[8..21]=B1_clase(14 val, mereu 0), rec[22..33]=B2_clase(12 val)
        #         B2_Clasa_15 = B2_total - sum(rec[22..33])
        # PRIORITATE față de is_sabre (sabre poate da fals-pozitiv pe aceste fișiere)
        def _detect_variant2b(av, rs):
            if rs != 34 or len(av) < rs * 5:
                return False
            for step in range(min(10, len(av) // rs)):
                base = step * rs
                if av[base] != 0:
                    return False
                if not (0 <= av[base + 4] <= 23):
                    return False
                # Discriminant față de Sabre: în variant2b banda 1 este defectă →
                # rec[5] (B1 total) = 0 și sum(rec[7:22]) (B1 clase) = 0 mereu.
                # În Sabre, rec[5] = B1 total > 0.
                if av[base + 5] != 0:
                    return False
            return av[4] > 0

        # ── Detecție variant4b totalizator: rec_size=10 în all_vals ──────────
        # Format: [h_next, B1, B2, B3, B4, B1dup, B2dup, B3dup, B4dup, 0]
        #         h_next = ora URMATOARE; datele sunt pentru ora h_next-1
        #         Funcționează și când n_lanes este raportat greșit ca 2
        def _detect_variant4b(av, nl):
            if nl not in (2, 4):
                return False
            # Verificăm 5 ore CONSECUTIVE la stride=10 pentru a evita fals-pozitive
            # pe fișiere 2-benzi standard unde o coincidență av[i+10]==(h+1)%24 apare accidental
            for i in range(min(20, len(av) - 60)):
                h = av[i]
                if 0 <= h <= 23 and 0 < av[i + 1] < 5000:
                    if all(i + s * 10 < len(av) and
                           av[i + s * 10] == (h + s) % 24
                           for s in range(1, 6)):
                        return True
            return False

        def _find_offset_4b(av):
            for i in range(min(20, len(av) - 60)):
                h = av[i]
                if 0 <= h <= 23 and 0 < av[i + 1] < 5000:
                    if all(i + s * 10 < len(av) and
                           av[i + s * 10] == (h + s) % 24
                           for s in range(1, 6)):
                        return i
            return 5

        # ── Detecție format 4 benzi cu preamble=11, h_end@0, totale@1-4 ──────
        # Format: preamble 11 val, apoi rec_size=66:
        #   rec[0]=h_end, rec[1..4]=Total_B1..B4, rec[5..19]=B1c1..15,
        #   rec[20..34]=B2c1..15, rec[35..49]=B3c1..15, rec[50..64]=B4c1..15
        # Detectare: hour_idx standard (7) dă mereu 0; idx 11 = oră secvențială
        def _detect_4band_hend(av, nl, rs):
            if nl != 4 or rs != 66 or len(av) < rs * 10:
                return False
            h7  = [av[step*66 + 7]  for step in range(10) if step*66 + 7  < len(av)]
            h11 = [av[step*66 + 11] for step in range(10) if step*66 + 11 < len(av)]
            zero_at_7 = all(h == 0 for h in h7)
            seq_at_11 = (all(0 <= h <= 23 for h in h11) and
                         all(h11[i+1] == (h11[i]+1) % 24 for i in range(len(h11)-1)))
            return zero_at_7 and seq_at_11

        # Ordinea de detecție contează: variant2b și variant4b au prioritate față de sabre
        is_variant2b   = (n_lanes == 2) and _detect_variant2b(all_vals, rec_size)
        is_variant4b   = _detect_variant4b(all_vals, n_lanes)
        is_4band_hend  = _detect_4band_hend(all_vals, n_lanes, rec_size)
        # Corecție critică: dacă e variant4b și n_lanes e raportat greșit ca 2, corectăm la 4
        if is_variant4b and n_lanes != 4:
            n_lanes = 4

        # ── Detecție totalizator v6: rec_size=6 în all_vals ──────────────────
        # Format: [h_end, B1, B2, B1dup, B2dup, 0x00] per record
        # h_end = ora de SFÂRŞIT a intervalului → ts = h_end - 1
        # B1dup == B1 şi B2dup == B2 (valori duplicate pentru verificare)
        # Folosit de contoarele setate totalizator (nu clasificator) cu 1 sau 2 benzi
        def _detect_tot_v6(av):
            if len(av) < 6 * 10:
                return False
            hours = [av[s * 6] for s in range(10) if s * 6 < len(av)]
            if not all(0 <= h <= 23 for h in hours):
                return False
            if not all(hours[i + 1] == (hours[i] + 1) % 24 for i in range(len(hours) - 1)):
                return False
            for step in range(10):
                rec = av[step * 6 : (step + 1) * 6]
                if rec[5] != 0:           return False   # terminator
                if rec[1] != rec[3]:      return False   # B1 == B1dup
                if rec[2] != rec[4]:      return False   # B2 == B2dup
            return True

        is_tot_v6 = (not is_variant2b) and (not is_variant4b) and (not is_4band_hend) and \
                    _detect_tot_v6(all_vals)

        is_sabre      = (not is_variant2b) and (not is_variant4b) and (not is_4band_hend) and \
                        (not is_tot_v6) and \
                        (is_sabre or _detect_sabre(all_vals, rec_size))
        is_totalizator = (n_lanes == 2) and (rec_size != 34) and (not is_variant2b) and \
                         (not is_tot_v6) and \
                         _detect_totalizator(raw, start_pos)


        # ── Parsing records ───────────────────────────────────────────────────
        rows = []

        if n_lanes == 6:
            # ── 6 benzi: format ADR 3000 Clasificator ────────────────────────
            # Structura în all_vals (stream varint pre-decodat):
            #   • Index block: primele 14 valori (se sare)
            #   • Per record: 98 valori fixe (8 header + 6 benzi × 15 clase)
            #     Header[0]   = necunoscut (0)
            #     Header[1]   = total L1 (control)
            #     Header[2]   = total L1? (duplicat) sau total general
            #     Header[3-7] = total L2..L6 (varint în raw, 1 valoare fiecare în all_vals)
            #   • Timestamp pornește din current_date (ora din raw[4])
            INDEX_BLOCK_6B = 14   # valori all_vals de sărit la început
            REC_SIZE_6B    = 98   # valori all_vals per record
            HEADER_6B      = 8    # valori header per record (înainte de clase)

            rec_idx = 0
            timestamp = current_date

            pos6 = INDEX_BLOCK_6B
            while pos6 + REC_SIZE_6B <= len(all_vals):
                rec = all_vals[pos6 : pos6 + REC_SIZE_6B]
                bands = []
                for lane in range(6):
                    s = HEADER_6B + lane * 15
                    bands.append(list(rec[s : s + 15]))

                rec_idx += 1
                row = {
                    "Contor":    site_id,
                    "Timestamp": timestamp,
                    "Data_Ora":  timestamp.strftime("%d.%m.%Y %H:%M"),
                    "N_Benzi":   6,
                }
                total_general = 0
                for b_idx, band in enumerate(bands, 1):
                    for cls_idx, val in enumerate(band, 1):
                        row[f"B{b_idx}_Clasa_{cls_idx}"] = val
                    tot_b = sum(band)
                    row[f"Total_B{b_idx}"] = tot_b
                    total_general += tot_b
                row["Total_General"] = total_general
                rows.append(row)

                pos6 += REC_SIZE_6B
                timestamp = timestamp + timedelta(hours=1)

        else:
            # ── 2 sau 4 benzi: parsing standard cu rec_size fix ───────────────

            # ── Totalizator: doar totaluri orare → Clasa_15 per bandă ─────────
            if is_totalizator:
                def _dv(raw, pos):
                    b = raw[pos]
                    if b >= 128 and pos + 1 < len(raw):
                        return (b & 0x7F) * 256 + raw[pos + 1], pos + 2
                    return b, pos + 1

                pos_t = start_pos
                while pos_t < len(raw):
                    h = raw[pos_t]
                    if h > 23:
                        pos_t += 1
                        continue
                    pos_t += 1
                    b1, pos_t = _dv(raw, pos_t)
                    b2, pos_t = _dv(raw, pos_t)
                    _, pos_t = _dv(raw, pos_t)  # B1 duplicat
                    _, pos_t = _dv(raw, pos_t)  # B2 duplicat
                    if pos_t < len(raw) and raw[pos_t] == 0:
                        pos_t += 1

                    if not rows:
                        timestamp = current_date.replace(hour=h, minute=0) - timedelta(hours=1)
                    else:
                        timestamp = rows[-1]["Timestamp"] + timedelta(hours=1)

                    # Toate clasele = 0, Clasa_15 = total orar
                    c_b1 = [0] * 15;
                    c_b1[14] = b1
                    c_b2 = [0] * 15;
                    c_b2[14] = b2

                    row = {
                        "Contor": site_id,
                        "Timestamp": timestamp,
                        "Data_Ora": timestamp.strftime("%d.%m.%Y %H:%M"),
                        "N_Benzi": 2,
                    }
                    for idx, val in enumerate(c_b1, 1): row[f"B1_Clasa_{idx}"] = val
                    row["Total_B1"] = b1
                    for idx, val in enumerate(c_b2, 1): row[f"B2_Clasa_{idx}"] = val
                    row["Total_B2"] = b2
                    for bn in range(3, 7):
                        for idx in range(1, 16): row[f"B{bn}_Clasa_{idx}"] = 0
                        row[f"Total_B{bn}"] = 0
                    row["Total_General"] = b1 + b2
                    rows.append(row)

            elif is_tot_v6:
                # ── Totalizator v6: rec_size=6, [h_end, B1, B2, B1dup, B2dup, 0x00] ──
                # h_end = ora de SFÂRŞIT → ts = h_end - 1
                # Totalurile B1/B2 merg în Clasa_15 (singura clasă disponibilă)
                # Rollover zi: când h_end scade cu mai mult de 12 faţă de precedentul
                tot_v6_date = current_date
                prev_h_end_v6 = None
                for step in range(len(all_vals) // 6):
                    rec = all_vals[step * 6 : (step + 1) * 6]
                    h_end = rec[0]
                    if h_end > 23:
                        continue
                    b1_tot = rec[1]
                    b2_tot = rec[2]
                    # Avansăm ziua la rollover
                    if prev_h_end_v6 is not None and h_end <= prev_h_end_v6 and \
                            (prev_h_end_v6 - h_end) > 12:
                        tot_v6_date += timedelta(days=1)
                    # ts = h_end - 1; rollover la h_end=0 → ts=23 pe ziua anterioară
                    h_ts = (h_end - 1) % 24
                    if h_end == 0:
                        ts = (tot_v6_date - timedelta(days=1)).replace(hour=23, minute=0)
                    else:
                        ts = tot_v6_date.replace(hour=h_ts, minute=0)
                    prev_h_end_v6 = h_end

                    row = {
                        "Contor":    site_id,
                        "Timestamp": ts,
                        "Data_Ora":  ts.strftime("%d.%m.%Y %H:%M"),
                        "N_Benzi":   max(n_lanes, 2),
                    }
                    total_general = 0
                    for b_idx, b_tot in enumerate([b1_tot, b2_tot], 1):
                        for idx in range(1, 16):
                            row[f"B{b_idx}_Clasa_{idx}"] = b_tot if idx == 15 else 0
                        row[f"Total_B{b_idx}"] = b_tot
                        total_general += b_tot
                    # Benzi suplimentare (dacă n_lanes > 2) = 0
                    for b_idx in range(3, max(n_lanes, 2) + 1):
                        for idx in range(1, 16): row[f"B{b_idx}_Clasa_{idx}"] = 0
                        row[f"Total_B{b_idx}"] = 0
                    row["Total_General"] = total_general
                    rows.append(row)

            elif is_variant2b:
                # ── Variant 2B: hour@4, B2_total@6, clase B2 la pos 22..33 ───
                # rec_size=34 în all_vals; rec[22..33]=clase B2 1-12
                # B2_Clasa_15 = B2_total - sum(clase 1-12)  (neidentificate)
                # B1 este banda defectă/neconectată → toate clasele = 0
                for step in range(len(all_vals) // 34):
                    base = step * 34
                    if base + 34 > len(all_vals):
                        break
                    rec      = all_vals[base : base + 34]
                    hour_val = rec[4]
                    if hour_val > 23:
                        continue
                    b2_total      = rec[6]
                    b2_classes_12 = list(rec[22:34])
                    b2_c15        = max(0, b2_total - sum(b2_classes_12))
                    b2_full = b2_classes_12 + [0, 0, b2_c15]
                    b1_full = [0] * 15
                    if not rows:
                        timestamp = current_date.replace(hour=hour_val, minute=0) - timedelta(hours=1)
                    else:
                        timestamp = rows[-1]["Timestamp"] + timedelta(hours=1)
                    row = {
                        "Contor":    site_id,
                        "Timestamp": timestamp,
                        "Data_Ora":  timestamp.strftime("%d.%m.%Y %H:%M"),
                        "N_Benzi":   2,
                    }
                    for cls_idx, val in enumerate(b1_full, 1): row[f"B1_Clasa_{cls_idx}"] = val
                    row["Total_B1"] = 0
                    for cls_idx, val in enumerate(b2_full, 1): row[f"B2_Clasa_{cls_idx}"] = val
                    row["Total_B2"] = b2_total
                    for bn in range(3, 7):
                        for cls_idx in range(1, 16): row[f"B{bn}_Clasa_{cls_idx}"] = 0
                        row[f"Total_B{bn}"] = 0
                    row["Total_General"] = b2_total
                    rows.append(row)

            elif is_variant4b:
                # ── Variant 4B totalizator: h_next@0, B1@1, B2@2, B3@3, B4@4 ─
                # rec_size=10 în all_vals; h_next=ora urmatoare, date pentru h_next-1
                # Toate totalurile merg la Clasa_15; funcționează și pt n_lanes=2 raportat greșit
                offset_4b  = _find_offset_4b(all_vals)
                cd_4b      = current_date
                prev_hn    = None
                for step in range((len(all_vals) - offset_4b) // 10):
                    base = offset_4b + step * 10
                    if base + 10 > len(all_vals):
                        break
                    rec    = all_vals[base : base + 10]
                    h_next = rec[0]
                    if h_next > 23:
                        continue
                    b1, b2, b3, b4 = rec[1], rec[2], rec[3], rec[4]
                    h_data = (h_next - 1) % 24
                    ts     = cd_4b.replace(hour=h_data)
                    if prev_hn is not None and h_next == 0:
                        cd_4b = cd_4b + timedelta(days=1)
                    prev_hn = h_next
                    c_b1 = [0]*15; c_b1[14] = b1
                    c_b2 = [0]*15; c_b2[14] = b2
                    c_b3 = [0]*15; c_b3[14] = b3
                    c_b4 = [0]*15; c_b4[14] = b4
                    row = {
                        "Contor":    site_id,
                        "Timestamp": ts,
                        "Data_Ora":  ts.strftime("%d.%m.%Y %H:%M"),
                        "N_Benzi":   4,
                    }
                    for cls_idx, val in enumerate(c_b1, 1): row[f"B1_Clasa_{cls_idx}"] = val
                    row["Total_B1"] = b1
                    for cls_idx, val in enumerate(c_b2, 1): row[f"B2_Clasa_{cls_idx}"] = val
                    row["Total_B2"] = b2
                    for cls_idx, val in enumerate(c_b3, 1): row[f"B3_Clasa_{cls_idx}"] = val
                    row["Total_B3"] = b3
                    for cls_idx, val in enumerate(c_b4, 1): row[f"B4_Clasa_{cls_idx}"] = val
                    row["Total_B4"] = b4
                    for bn in range(5, 7):
                        for cls_idx in range(1, 16): row[f"B{bn}_Clasa_{cls_idx}"] = 0
                        row[f"Total_B{bn}"] = 0
                    row["Total_General"] = b1 + b2 + b3 + b4
                    rows.append(row)

            elif is_4band_hend:
                # ── Format 4B cu preamble=11, h_end@rec[0], totale@rec[1..4] ─
                # rec[5..19]=B1c1..15, rec[20..34]=B2, rec[35..49]=B3, rec[50..64]=B4
                # h_end = ora de SFARSIT a intervalului; timestamp = h_end - 1
                PREAMBLE_4H = 11
                cd_4h = datetime(raw[7] + 2000, raw[6], raw[5], 0, 0)
                prev_hend = None
                n_recs_4h = (len(all_vals) - PREAMBLE_4H) // rec_size
                for step in range(n_recs_4h):
                    base = PREAMBLE_4H + step * rec_size
                    if base + rec_size > len(all_vals):
                        break
                    rec = all_vals[base : base + rec_size]
                    h_end = rec[0]
                    if h_end > 23:
                        continue
                    h_data = (h_end - 1) % 24
                    ts = cd_4h.replace(hour=h_data)         # timestamp ÎNAINTE de avansarea zilei
                    if prev_hend is not None and h_end == 0:
                        cd_4h = cd_4h + timedelta(days=1)   # avansăm ziua DUPĂ ce am setat ts
                    prev_hend = h_end
                    b1 = list(rec[5:20])
                    b2 = list(rec[20:35])
                    b3 = list(rec[35:50])
                    b4 = list(rec[50:65])
                    row = {
                        "Contor":    site_id,
                        "Timestamp": ts,
                        "Data_Ora":  ts.strftime("%d.%m.%Y %H:%M"),
                        "N_Benzi":   4,
                    }
                    for cls_idx, val in enumerate(b1, 1): row[f"B1_Clasa_{cls_idx}"] = val
                    row["Total_B1"] = sum(b1)
                    for cls_idx, val in enumerate(b2, 1): row[f"B2_Clasa_{cls_idx}"] = val
                    row["Total_B2"] = sum(b2)
                    for cls_idx, val in enumerate(b3, 1): row[f"B3_Clasa_{cls_idx}"] = val
                    row["Total_B3"] = sum(b3)
                    for cls_idx, val in enumerate(b4, 1): row[f"B4_Clasa_{cls_idx}"] = val
                    row["Total_B4"] = sum(b4)
                    for bn in range(5, 7):
                        for cls_idx in range(1, 16): row[f"B{bn}_Clasa_{cls_idx}"] = 0
                        row[f"Total_B{bn}"] = 0
                    row["Total_General"] = sum(b1) + sum(b2) + sum(b3) + sum(b4)
                    rows.append(row)

            else:
                # parsare normală cu rec_size fix
                n_records = len(all_vals) // rec_size
                for step in range(n_records):


                    base = step * rec_size
                    if base + rec_size > len(all_vals):
                        break

                    rec = all_vals[base : base + rec_size]

                    # Ora
                    hour_val = rec[4] if is_sabre else rec[hour_idx]
                    if hour_val > 23:
                        continue

                    # Extragere clase per bandă
                    if is_sabre:
                        # B1: rec[5]=total, rec[7:22]=clase 1-15 (15 valori complete)
                        # B2: rec[6]=total, rec[22:34]=clase 1-12 (12 valori);
                        #     clasele 13-14=0, clasa 15=total-sum(clase 1-12) (vehicule neidentificate)
                        b2_total      = rec[6]
                        b2_classes_12 = list(rec[22:34])
                        b2_c15        = max(0, b2_total - sum(b2_classes_12))
                        bands = [list(rec[7:22]),
                                 b2_classes_12 + [0, 0, b2_c15]]
                        bands += [[0] * 15] * (n_lanes - 2)
                    else:
                        bands = []
                        for lane in range(n_lanes):
                            s = lane_start + lane * 15

                            if has_totals_prefix and lane == n_lanes - 1:
                                cls_current = rec[s : s + 9]
                                next_base   = (step + 1) * rec_size
                                cls_next    = (all_vals[next_base : next_base + 6]
                                               if next_base + 6 <= len(all_vals) else [])
                                raw_band = list(cls_current) + list(cls_next)
                            else:
                                raw_band = list(rec[s : s + 15])

                            raw_band = raw_band + [0] * (15 - len(raw_band))
                            bands.append(raw_band)

                    while len(bands) < 4:
                        bands.append([0] * 15)

                    if not rows:
                        timestamp = current_date.replace(hour=hour_val, minute=0) - timedelta(hours=1)
                    else:
                        timestamp = rows[-1]["Timestamp"] + timedelta(hours=1)

                    row = {
                        "Contor":    site_id,
                        "Timestamp": timestamp,
                        "Data_Ora":  timestamp.strftime("%d.%m.%Y %H:%M"),
                        "N_Benzi":   n_lanes,
                    }

                    total_general = 0
                    for b_idx, band in enumerate(bands[:n_lanes], 1):
                        for cls_idx, val in enumerate(band, 1):
                            row[f"B{b_idx}_Clasa_{cls_idx}"] = val
                        tot_b = sum(band)
                        row[f"Total_B{b_idx}"] = tot_b
                        total_general += tot_b

                    max_bands = 6
                    for b_idx in range(n_lanes + 1, max_bands + 1):
                        for cls_idx in range(1, 16):
                            row[f"B{b_idx}_Clasa_{cls_idx}"] = 0
                        row[f"Total_B{b_idx}"] = 0

                    row["Total_General"] = total_general
                    rows.append(row)

        df = pd.DataFrame(rows)

    # ── Finalizare DataFrame ──────────────────────────────────────────────────
    if df.empty:
        return df, site_id, current_date, n_lanes

    df["Timestamp"] = pd.to_datetime(df["Timestamp"])
    df = df.set_index("Timestamp")
    df = df[~df.index.duplicated(keep="last")]
    full_range = pd.date_range(start=df.index.min(), end=df.index.max(), freq="h")
    df = df.reindex(full_range)
    df["Contor"]  = df["Contor"].fillna(site_id)
    df["N_Benzi"] = df["N_Benzi"].fillna(n_lanes).astype(int)
    numeric_cols  = df.select_dtypes(include="number").columns
    df[numeric_cols] = df[numeric_cols].fillna(0)
    df["Data_Ora"] = df.index.strftime("%d.%m.%Y %H:%M")
    df = df.reset_index(drop=True)

    return df, site_id, current_date, n_lanes
# ══════════════════════════════════════════════════════════════════════════════
# FORMATARE ȘI GRAFICE EXCEL
# ══════════════════════════════════════════════════════════════════════════════


def quick_scan_bin(filepath):
    """
    Citește rapid un fișier .bin și returnează (site_id, n_records, n_lanes).
    Detectează automat RO23 (PEEK/Sabre), Ro04R și numărul de benzi.
    """
    try:
        with open(filepath, 'rb') as f:
            data = f.read()

        # site_id din header binar
        try:
            header_text = "".join([chr(b) if 32 <= b <= 126 else " " for b in data[0:150]])
            matches = re.findall(r'000+(\d{4})', header_text)
            site_id = matches[-1] if matches else "????"
        except Exception:
            site_id = "????"

        is_ro04r = (data.find(b'Ro04R') != -1) and (data.find(b'RO23') == -1)
        n_records = 0

        # Detectare număr benzi din header RO23
        pos_ro23 = data.find(b'RO23')
        if pos_ro23 >= 0 and not is_ro04r:
            raw_lanes = data[pos_ro23 + 11] // 2
            n_lanes   = raw_lanes if raw_lanes in (2, 4, 6) else 2
        else:
            n_lanes = 2

        # Dimensiune record în funcție de benzi
        rec_size = {2: 34, 4: 66, 6: 98}.get(n_lanes, 34)

        if is_ro04r:
            ro04r_idx = data.find(b'Ro04R')
            pos = ro04r_idx + 14
            while pos < len(data):
                b = data[pos]
                if b == 0x00:
                    pos += 1; continue
                if b > 23:
                    pos += 1; continue
                pos += 1
                i = pos
                ok = True
                for _ in range(4):
                    if i >= len(data):
                        ok = False; break
                    bv = data[i]
                    i += 2 if bv >= 128 else 1
                if ok:
                    n_records += 1
                    pos = i
                else:
                    break
        else:
            if pos_ro23 == -1:
                return site_id, 0, n_lanes
            # Decodăm în all_vals și numărăm recorduri
            start = pos_ro23 + 68
            all_vals = []
            i = start
            while i < len(data):
                b = data[i]
                if b >= 128:
                    if i + 1 < len(data):
                        all_vals.append((b & 0x7F) * 256 + data[i + 1])
                        i += 2
                    else: break
                else:
                    all_vals.append(b)
                    i += 1
            n_records = len(all_vals) // rec_size

        return site_id, n_records, n_lanes
    except Exception:
        return "????", 0, 2




def process_multiple_files(filepaths, output_dir=None, stop_event=None):
    contoare_data  = {}   # {site_id: [df1, df2, ...]}
    contoare_lanes = {}   # {site_id: n_lanes}

    for filepath in filepaths:
        if stop_event and stop_event.is_set():
            return None   # anulat de utilizator
        try:
            df, site_id, start_date, n_lanes = process_peek_bin(filepath)
            if df is None or df.empty:
                continue
            if site_id not in contoare_data:
                contoare_data[site_id]  = []
                contoare_lanes[site_id] = n_lanes
            contoare_data[site_id].append(df)
            # Păstrăm numărul maxim de benzi găsite pentru acest contor
            if n_lanes > contoare_lanes[site_id]:
                contoare_lanes[site_id] = n_lanes
        except Exception as e:
            print(f"Sărit fișier corupt {filepath}: {e}")
            continue

    if not contoare_data:
        print("Nu s-au găsit date valide în niciun fișier.")
        return None

    rezultate_finale = []
    out_dir = os.path.dirname(os.path.abspath(filepaths[0]))

    for site_id, lista_dfs in contoare_data.items():
        n_lanes = contoare_lanes[site_id]

        df_total = pd.concat(lista_dfs, ignore_index=True)

        # Convertim Data_Ora în datetime real
        df_total["Data_Ora_dt"] = pd.to_datetime(
            df_total["Data_Ora"],
            format="%d.%m.%Y %H:%M",
            errors="coerce"
        )

        # Sortare cronologică
        df_total = df_total.sort_values(["Contor", "Data_Ora_dt"])

        # Eliminare duplicate pentru același contor + aceeași oră
        df_total = df_total.drop_duplicates(
            subset=["Contor", "Data_Ora_dt"],
            keep="last"
        )

        # Ștergem coloanele auxiliare
        df_total = df_total.drop(columns=["Data_Ora_dt"], errors="ignore")

        # Eliminăm benzile cu total zero dacă sunt mai puține benzi reale
        # (benzile 0 au fost adăugate pentru compatibilitate, nu le exportăm)
        # Păstrăm doar benzile până la n_lanes
        cols_to_drop = []
        for b in range(n_lanes + 1, 7):
            for c in range(1, 16):
                col = f"B{b}_Clasa_{c}"
                if col in df_total.columns:
                    cols_to_drop.append(col)
            if f"Total_B{b}" in df_total.columns:
                cols_to_drop.append(f"Total_B{b}")
        if cols_to_drop:
            df_total = df_total.drop(columns=cols_to_drop)

        # Recalculăm Total_General după tăierea benzilor
        band_tot_cols = [f"Total_B{b}" for b in range(1, n_lanes + 1)
                         if f"Total_B{b}" in df_total.columns]
        df_total["Total_General"] = df_total[band_tot_cols].sum(axis=1)

        # Eliminăm coloana N_Benzi din export (internă)
        df_export = df_total.drop(columns=["N_Benzi"], errors="ignore")
        df_export = df_export.reset_index(drop=True)

        output_fn = os.path.join(out_dir, f"Raport_Clase_Peek_{site_id}.xlsx")

        # Suma totală pe toate benzile
        all_band_totals = sum(df_export[c].sum() for c in band_tot_cols
                              if c in df_export.columns)

        # Salvăm Excel-ul brut
        df_export.to_excel(output_fn, index=False, sheet_name="Date Detaliate")

        # Adăugăm graficele DOAR dacă există date
        if all_band_totals > 0:
            add_charts_and_formatting(output_fn, df_export, site_id)
        else:
            print(f"Atenție: Contorul {site_id} are trafic zero. Excel generat fără grafice.")

        suma_b1 = df_export['Total_B1'].sum() if 'Total_B1' in df_export.columns else 0
        suma_b2 = df_export['Total_B2'].sum() if 'Total_B2' in df_export.columns else 0

        rezultate_finale.append({
            'path': output_fn, 'id': site_id, 'randuri': len(df_export),
            'b1': suma_b1, 'b2': suma_b2, 'n_lanes': n_lanes
        })

    if not rezultate_finale:
        return None

    return rezultate_finale


