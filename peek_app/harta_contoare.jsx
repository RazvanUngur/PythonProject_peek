// harta_contoare.jsx — Hartă interactivă CESTRIN
// Integrare: from harta_api import register_harta_routes → register_harta_routes(app)

import { useState, useEffect, useRef, useCallback } from "react";

// ── Status config ─────────────────────────────────────────────────────────────
const STATUS_CFG = {
  CLASIFICATOR: { label: "Clasificator",     color: "#16a34a", fill: "#22c55e", ring: "#15803d" },
  TOTALIZATOR:  { label: "Totalizator",      color: "#b45309", fill: "#f59e0b", ring: "#d97706" },
  FARA_COM:     { label: "Fără comunicare",  color: "#7c3aed", fill: "#a78bfa", ring: "#6d28d9" },
  DEFECT:       { label: "Defect",           color: "#b91c1c", fill: "#ef4444", ring: "#dc2626" },
  NELOCALIZAT:  { label: "Nelocalizat",      color: "#374151", fill: "#9ca3af", ring: "#6b7280" },
};

// ── Leaflet loader ────────────────────────────────────────────────────────────
function useLeaflet() {
  const [ready, setReady] = useState(!!window.L);
  useEffect(() => {
    if (window.L) { setReady(true); return; }
    const css = document.createElement("link");
    css.rel = "stylesheet";
    css.href = "https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.9.4/leaflet.min.css";
    document.head.appendChild(css);
    const script = document.createElement("script");
    script.src = "https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.9.4/leaflet.min.js";
    script.onload = () => setReady(true);
    document.head.appendChild(script);
  }, []);
  return ready;
}

// ── SVG marker icon ───────────────────────────────────────────────────────────
function makeIcon(L, status, opts = {}) {
  const { selected = false, draggable = false } = opts;
  const s    = STATUS_CFG[status] || STATUS_CFG.NELOCALIZAT;
  const r    = selected ? 13 : 10;
  const ring = selected ? 3.5 : 2;
  const tot  = (r + ring) * 2;
  const cx   = tot / 2;

  const inner = draggable
    ? `<circle cx="${cx}" cy="${cx}" r="${r * 0.38}" fill="white" fill-opacity="0.95"/>
       <line x1="${cx}" y1="${cx - r*0.55}" x2="${cx}" y2="${cx + r*0.55}" stroke="white" stroke-width="1.5" stroke-linecap="round"/>
       <line x1="${cx - r*0.55}" y1="${cx}" x2="${cx + r*0.55}" y2="${cx}" stroke="white" stroke-width="1.5" stroke-linecap="round"/>`
    : selected
      ? `<circle cx="${cx}" cy="${cx}" r="${r * 0.38}" fill="white" fill-opacity="0.9"/>`
      : "";

  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${tot}" height="${tot}">
    <circle cx="${cx}" cy="${cx}" r="${r}" fill="${s.fill}" stroke="${s.ring}" stroke-width="${ring}" fill-opacity="${selected ? 1 : 0.9}"/>
    ${inner}
  </svg>`;

  return L.divIcon({
    html: svg, className: "",
    iconSize:   [tot, tot],
    iconAnchor: [cx, cx],
    popupAnchor:[0, -(cx + 4)],
  });
}

// ── Formatare numere ──────────────────────────────────────────────────────────
const fmt = v => (v == null ? "—" : Number(v).toLocaleString("ro-RO"));

// ── Popup HTML ────────────────────────────────────────────────────────────────
function buildPopup(c) {
  const s    = STATUS_CFG[c.status] || STATUS_CFG.NELOCALIZAT;
  const mzlStr = c.mzl != null ? `${fmt(c.mzl)} veh/zi` : "Indisponibil";
  const manualBadge = c.mzl_sursa === "manual"
    ? `<span style="font-size:10px;background:#dbeafe;color:#1d4ed8;padding:1px 6px;border-radius:8px;font-weight:600;margin-left:5px">manual</span>`
    : "";
  const cls15str = (c.total_cls15 != null && c.total_veh)
    ? `${fmt(c.total_cls15)} (${Math.round(c.total_cls15 / c.total_veh * 100)}%)`
    : "—";

  return `
    <div style="font-family:'Segoe UI',system-ui,sans-serif;min-width:230px;padding:4px 2px">
      <div style="display:flex;align-items:center;gap:8px;margin-bottom:10px">
        <span style="width:11px;height:11px;border-radius:50%;background:${s.fill};flex-shrink:0;display:inline-block"></span>
        <span style="font-weight:700;font-size:15px;color:#0f172a;letter-spacing:-0.3px">${c.contor}</span>
      </div>
      <table style="width:100%;border-collapse:collapse;font-size:12.5px">
        <tr><td style="color:#64748b;padding:3px 0;width:115px">Drum</td>
            <td style="font-weight:600;color:#1e293b">${c.drum || "—"}</td></tr>
        <tr><td style="color:#64748b;padding:3px 0">Poziție km</td>
            <td style="font-weight:600;color:#1e293b">${c.pozitie_km || "—"}</td></tr>
        <tr><td style="color:#64748b;padding:3px 0">Localitate</td>
            <td style="font-weight:600;color:#1e293b">${c.localitate || "—"}</td></tr>
        <tr><td style="color:#64748b;padding:3px 0">Tip contor</td>
            <td style="font-weight:600;color:#1e293b">${c.tip || "—"}</td></tr>
        <tr><td colspan="2"><hr style="border:none;border-top:1px solid #e2e8f0;margin:6px 0"/></td></tr>
        <tr><td style="color:#64748b;padding:3px 0">MZL ${c.luna_ref || ""}</td>
            <td style="font-weight:700;color:${s.fill};font-size:13.5px">${mzlStr}${manualBadge}</td></tr>
        <tr><td style="color:#64748b;padding:3px 0">Clasa 15</td>
            <td style="color:#475569">${cls15str}</td></tr>
        <tr><td style="color:#64748b;padding:3px 0">Acoperire ore</td>
            <td style="color:#475569">${c.acop_pct != null ? c.acop_pct + "%" : "—"}</td></tr>
        ${c.ultima_luna && c.status !== "CLASIFICATOR" && c.status !== "TOTALIZATOR"
          ? `<tr><td style="color:#64748b;padding:3px 0">Ultimele date</td>
                 <td style="color:#475569">${c.ultima_luna}</td></tr>` : ""}
      </table>
      <div style="margin-top:9px;padding:5px 10px;border-radius:6px;
                  background:${s.fill}18;border:1px solid ${s.ring};
                  display:inline-flex;align-items:center;gap:6px">
        <span style="width:8px;height:8px;border-radius:50%;background:${s.fill};display:inline-block"></span>
        <span style="font-size:11.5px;font-weight:600;color:${s.color}">${s.label}</span>
      </div>
      ${c.status === "NELOCALIZAT"
        ? `<div style="margin-top:6px;font-size:11px;color:#7c3aed">
             📍 Trage markerul pe poziția corectă pentru a salva coordonatele.
           </div>` : ""}
    </div>`;
}

// ── Date mock ─────────────────────────────────────────────────────────────────
const MOCK = [
  { contor:"DN1_001",  drum:"DN1",  pozitie_km:"125+400", localitate:"Comarnic",  tip:"PEEK", lat:45.254, lng:25.639, mzl:8420,  mzl_sursa:"calculat", luna_ref:"04.2026", status:"CLASIFICATOR", acop_pct:94,  total_veh:252600, total_cls15:8000,  ultima_luna:"04.2026" },
  { contor:"DN1_002",  drum:"DN1",  pozitie_km:"158+200", localitate:"Sinaia",    tip:"PEEK", lat:45.348, lng:25.546, mzl:6180,  mzl_sursa:"manual",   luna_ref:"04.2026", status:"CLASSIFICATOR",acop_pct:88,  total_veh:185400, total_cls15:5500,  ultima_luna:"04.2026" },
  { contor:"DN2_005",  drum:"DN2",  pozitie_km:"42+000",  localitate:"Urziceni",  tip:"VEK",  lat:44.717, lng:26.641, mzl:4320,  mzl_sursa:"calculat", luna_ref:"04.2026", status:"TOTALIZATOR",  acop_pct:55,  total_veh:129600, total_cls15:19000, ultima_luna:"04.2026" },
  { contor:"DN7_003",  drum:"DN7",  pozitie_km:"87+600",  localitate:"Câmpulung", tip:"PEEK", lat:45.272, lng:25.048, mzl:0,     mzl_sursa:"calculat", luna_ref:"04.2026", status:"DEFECT",       acop_pct:0,   total_veh:0,      total_cls15:0,     ultima_luna:"03.2026" },
  { contor:"A1_010",   drum:"A1",   pozitie_km:"KM 210",  localitate:"Pitești",   tip:"PEEK", lat:44.857, lng:24.869, mzl:18200, mzl_sursa:"manual",   luna_ref:"04.2026", status:"CLASIFICATOR", acop_pct:99,  total_veh:546000, total_cls15:15000, ultima_luna:"04.2026" },
  { contor:"DN13_002", drum:"DN13", pozitie_km:"22+100",  localitate:"Brașov",    tip:"VEK",  lat:45.650, lng:25.607, mzl:null,  mzl_sursa:null,       luna_ref:"04.2026", status:"FARA_COM",     acop_pct:0,   total_veh:null,   total_cls15:null,  ultima_luna:"02.2026" },
  { contor:"DN15_007", drum:"DN15", pozitie_km:"215+000", localitate:"Reghin",    tip:"PEEK", lat:46.772, lng:24.918, mzl:3800,  mzl_sursa:"calculat", luna_ref:"04.2026", status:"CLASIFICATOR", acop_pct:78,  total_veh:114000, total_cls15:9000,  ultima_luna:"04.2026" },
  { contor:"DN18_001", drum:"DN18", pozitie_km:"7+200",   localitate:"Baia Mare", tip:"VEK",  lat:47.658, lng:23.570, mzl:5600,  mzl_sursa:"calculat", luna_ref:"04.2026", status:"TOTALIZATOR",  acop_pct:40,  total_veh:168000, total_cls15:22000, ultima_luna:"04.2026" },
  { contor:"DN22_003", drum:"DN22", pozitie_km:"64+800",  localitate:"Tulcea",    tip:"PEEK", lat:45.170, lng:28.803, mzl:1200,  mzl_sursa:"calculat", luna_ref:"04.2026", status:"CLASIFICATOR", acop_pct:68,  total_veh:36000,  total_cls15:2000,  ultima_luna:"04.2026" },
  { contor:"DN65_001", drum:"DN65", pozitie_km:"0+000",   localitate:"Craiova",   tip:"PEEK", lat:null,   lng:null,   mzl:null,  mzl_sursa:null,       luna_ref:"04.2026", status:"NELOCALIZAT",  acop_pct:null,total_veh:null,   total_cls15:null,  ultima_luna:"04.2026" },
  { contor:"DN6_008",  drum:"DN6",  pozitie_km:"333+000", localitate:"Timișoara", tip:"VEK",  lat:45.749, lng:21.209, mzl:9100,  mzl_sursa:"calculat", luna_ref:"04.2026", status:"CLASIFICATOR", acop_pct:91,  total_veh:273000, total_cls15:8000,  ultima_luna:"04.2026" },
];

// ═══════════════════════════════════════════════════════════════════════════════
export default function HartaContoare() {
  const leafletReady  = useLeaflet();
  const mapRef        = useRef(null);
  const mapInst       = useRef(null);
  const markersRef    = useRef({});       // { contor: L.Marker }
  const dragMarkerRef = useRef(null);     // marker drag nelocalizat activ

  const [contoare,   setContoare]   = useState([]);
  const [loading,    setLoading]    = useState(true);
  const [error,      setError]      = useState(null);
  const [selected,   setSelected]   = useState(null);
  const [saving,     setSaving]     = useState(null); // contor în curs de salvare
  const [savedOk,    setSavedOk]    = useState(null);
  const [filterSt,   setFilterSt]   = useState("ALL");
  const [filterTip,  setFilterTip]  = useState("ALL");
  const [searchQ,    setSearchQ]    = useState("");
  const [dataSource, setDataSource] = useState("api");
  const [lunaRef,    setLunaRef]    = useState("");

  // ── Load date ───────────────────────────────────────────────────────────────
  const loadData = useCallback(async (src) => {
    setLoading(true); setError(null);
    try {
      if (src === "mock") {
        await new Promise(r => setTimeout(r, 300));
        setContoare(MOCK);
        setLunaRef(MOCK[0]?.luna_ref || "");
      } else {
        const res = await fetch("/api/harta_contoare");
        if (!res.ok) throw new Error(`HTTP ${res.status}`);
        const json = await res.json();
        if (json.error) throw new Error(json.error);
        setContoare(json);
        setLunaRef(json.find(c => c.luna_ref)?.luna_ref || "");
      }
    } catch (e) {
      if (src === "api") {
        setDataSource("mock"); setContoare(MOCK);
        setLunaRef(MOCK[0]?.luna_ref || "");
        setError("API indisponibil — date demonstrative");
      } else {
        setError(e.message);
      }
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => { loadData(dataSource); }, [dataSource]);

  // ── Init hartă ──────────────────────────────────────────────────────────────
  useEffect(() => {
    if (!leafletReady || !mapRef.current || mapInst.current) return;
    const L = window.L;
    const map = L.map(mapRef.current, {
      center: [45.9432, 24.9668], zoom: 7, zoomControl: false,
    });
    L.control.zoom({ position: "bottomright" }).addTo(map);
    L.tileLayer(
      "https://{s}.basemaps.cartocdn.com/light_all/{z}/{x}/{y}{r}.png",
      { attribution: "© OpenStreetMap © CARTO", maxZoom: 19 }
    ).addTo(map);
    mapInst.current = map;
    return () => { map.remove(); mapInst.current = null; };
  }, [leafletReady]);

  // ── Filtre ──────────────────────────────────────────────────────────────────
  const filtered = contoare.filter(c => {
    if (filterSt !== "ALL" && c.status !== filterSt) return false;
    if (filterTip !== "ALL" && c.tip !== filterTip)  return false;
    if (searchQ) {
      const q = searchQ.toLowerCase();
      return [c.contor, c.drum, c.localitate, c.pozitie_km]
        .some(v => v && v.toLowerCase().includes(q));
    }
    return true;
  });

  // ── Salvare coordonate ───────────────────────────────────────────────────────
  const saveCoords = useCallback(async (contor, lat, lng) => {
    setSaving(contor);
    try {
      if (dataSource === "mock") {
        // Mock: actualizăm state local
        await new Promise(r => setTimeout(r, 400));
        setContoare(prev => prev.map(c =>
          c.contor === contor
            ? { ...c, lat, lng, status: c.status === "NELOCALIZAT" ? "FARA_COM" : c.status }
            : c
        ));
      } else {
        const res = await fetch("/api/harta_contoare/coordonate", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ contor, lat, lng }),
        });
        const json = await res.json();
        if (!res.ok || json.error) throw new Error(json.error || "Eroare server");
        // Reîncărcăm datele pentru a reflecta noul status
        await loadData(dataSource);
      }
      setSavedOk(contor);
      setTimeout(() => setSavedOk(null), 3000);
    } catch (e) {
      setError(`Eroare salvare coordonate: ${e.message}`);
    } finally {
      setSaving(null);
    }
  }, [dataSource, loadData]);

  // ── Render markere pe hartă ─────────────────────────────────────────────────
  useEffect(() => {
    const L = window.L;
    const map = mapInst.current;
    if (!L || !map || contoare.length === 0) return;

    // Curăță markere vechi
    Object.values(markersRef.current).forEach(m => m.remove());
    markersRef.current = {};
    if (dragMarkerRef.current) { dragMarkerRef.current.remove(); dragMarkerRef.current = null; }

    const located    = filtered.filter(c => c.lat && c.lng);
    const unlocated  = filtered.filter(c => !c.lat || !c.lng);

    // Markere pentru contoare cu coordonate
    located.forEach(c => {
      const isSel = selected === c.contor;
      const marker = L.marker([c.lat, c.lng], {
        icon: makeIcon(L, c.status, { selected: isSel }),
        zIndexOffset: isSel ? 1000 : 0,
        title: c.contor,
      });
      marker.bindPopup(buildPopup(c), {
        maxWidth: 290, className: "cestrin-popup",
      });
      marker.on("click", () => setSelected(c.contor));
      marker.on("popupclose", () => setSelected(null));
      marker.addTo(map);
      markersRef.current[c.contor] = marker;
    });

    // Markere draggable pentru contoare fără coordonate
    // Le plasăm în centrul României, ușor offset pe y ca să nu se suprapună
    unlocated.forEach((c, i) => {
      const lat = 45.9432 + (i - unlocated.length / 2) * 0.3;
      const lng = 24.9668;
      const marker = L.marker([lat, lng], {
        icon: makeIcon(L, "NELOCALIZAT", { draggable: true }),
        draggable: true,
        title: `${c.contor} — trage pe poziția corectă`,
        zIndexOffset: 500,
      });
      marker.bindPopup(buildPopup(c), { maxWidth: 290, className: "cestrin-popup" });
      marker.on("click", () => { setSelected(c.contor); marker.openPopup(); });

      marker.on("dragend", (e) => {
        const { lat: newLat, lng: newLng } = e.target.getLatLng();
        // Afișăm popup de confirmare
        const popupHtml = `
          <div style="font-family:'Segoe UI',system-ui,sans-serif;padding:4px">
            <div style="font-weight:700;margin-bottom:8px;color:#0f172a">${c.contor}</div>
            <div style="font-size:12px;color:#64748b;margin-bottom:10px">
              Lat: <b>${newLat.toFixed(5)}</b><br/>
              Lng: <b>${newLng.toFixed(5)}</b>
            </div>
            <button id="btn-save-${c.contor.replace(/[^a-z0-9]/gi,"_")}"
              style="background:#16a34a;color:white;border:none;padding:6px 14px;
                     border-radius:6px;cursor:pointer;font-weight:600;font-size:12px;width:100%">
              ✓ Salvează coordonatele
            </button>
          </div>`;
        marker.bindPopup(popupHtml, { maxWidth: 230, className: "cestrin-popup" })
              .openPopup();

        // Atașăm listener pe buton după ce popup-ul e în DOM
        setTimeout(() => {
          const btn = document.getElementById(
            `btn-save-${c.contor.replace(/[^a-z0-9]/gi,"_")}`
          );
          if (btn) {
            btn.onclick = () => {
              marker.closePopup();
              saveCoords(c.contor, newLat, newLng);
            };
          }
        }, 100);
      });

      marker.addTo(map);
    });

  }, [filtered, selected, saveCoords]);

  // ── Focus pe card click ──────────────────────────────────────────────────────
  const focusContor = useCallback((c) => {
    const map = mapInst.current;
    if (!map) return;
    if (c.lat && c.lng) {
      map.setView([c.lat, c.lng], 12, { animate: true, duration: 0.6 });
      const m = markersRef.current[c.contor];
      if (m) { m.openPopup(); setSelected(c.contor); }
    }
  }, []);

  // ── Popup CSS ─────────────────────────────────────────────────────────────────
  useEffect(() => {
    const style = document.createElement("style");
    style.textContent = `
      .cestrin-popup .leaflet-popup-content-wrapper {
        border-radius:10px!important;box-shadow:0 8px 32px rgba(0,0,0,0.16)!important;
        padding:0!important;border:1px solid #e2e8f0;
      }
      .cestrin-popup .leaflet-popup-content { margin:14px 16px!important; }
      .cestrin-popup .leaflet-popup-tip { background:white!important; }
    `;
    document.head.appendChild(style);
    return () => style.remove();
  }, []);

  // ── Stats ─────────────────────────────────────────────────────────────────────
  const tipuri = [...new Set(contoare.map(c => c.tip).filter(Boolean))].sort();
  const statCount = (st) => contoare.filter(c => c.status === st).length;

  // ── Render ────────────────────────────────────────────────────────────────────
  return (
    <div style={{
      fontFamily:"'Segoe UI',system-ui,sans-serif",
      display:"flex", flexDirection:"column",
      height:"100vh", background:"#f1f5f9", overflow:"hidden",
    }}>

      {/* ── HEADER ─────────────────────────────────────────────────────────── */}
      <div style={{
        background:"linear-gradient(135deg,#0f172a 0%,#1e3a5f 100%)",
        padding:"11px 18px", display:"flex", alignItems:"center", gap:14,
        boxShadow:"0 2px 12px rgba(0,0,0,0.3)", flexShrink:0,
      }}>
        <div style={{
          width:34,height:34,borderRadius:8,
          background:"linear-gradient(135deg,#3b82f6,#1d4ed8)",
          display:"flex",alignItems:"center",justifyContent:"center",fontSize:17,
        }}>🗺️</div>
        <div>
          <div style={{color:"#fff",fontWeight:700,fontSize:15.5,letterSpacing:-0.3}}>
            Hartă Contoare Trafic
          </div>
          <div style={{color:"#94a3b8",fontSize:11}}>
            CESTRIN · Referință: {lunaRef || "—"}
            {dataSource === "mock" && <span style={{color:"#fbbf24",marginLeft:8}}>⚠ Demo</span>}
          </div>
        </div>
        <div style={{flex:1}}/>

        {/* Stat pills */}
        {!loading && Object.entries(STATUS_CFG).map(([k, v]) => {
          const n = statCount(k);
          if (n === 0) return null;
          return (
            <div key={k} onClick={() => setFilterSt(f => f === k ? "ALL" : k)}
              style={{
                padding:"4px 10px",borderRadius:20,cursor:"pointer",
                background: filterSt === k ? v.fill + "33" : "rgba(255,255,255,0.07)",
                color:v.fill, fontSize:11.5, fontWeight:600,
                border:`1px solid ${filterSt === k ? v.ring : "rgba(255,255,255,0.12)"}`,
                transition:"all .15s",
              }}
            >{n} {v.label}</div>
          );
        })}

        <button onClick={() => loadData(dataSource)} style={{
          background:"rgba(255,255,255,0.1)",border:"1px solid rgba(255,255,255,0.2)",
          color:"#fff",borderRadius:7,padding:"5px 12px",cursor:"pointer",
          fontSize:12,fontWeight:600,
        }}>↺ Actualizează</button>

        <button onClick={() => setDataSource(s => s==="api"?"mock":"api")} style={{
          background: dataSource==="mock" ? "rgba(245,158,11,0.15)" : "rgba(255,255,255,0.07)",
          border:`1px solid ${dataSource==="mock" ? "#d97706" : "rgba(255,255,255,0.15)"}`,
          color: dataSource==="mock" ? "#fbbf24" : "#64748b",
          borderRadius:7,padding:"5px 11px",cursor:"pointer",fontSize:12,fontWeight:600,
        }}>{dataSource==="mock" ? "Demo" : "Live"}</button>
      </div>

      {/* ── BANNER eroare / succes ─────────────────────────────────────────── */}
      {error && (
        <div style={{background:"#fef3c7",borderBottom:"1px solid #fcd34d",
          padding:"7px 18px",fontSize:12.5,color:"#92400e",display:"flex",
          alignItems:"center",gap:8,flexShrink:0}}>
          ⚠️ {error}
          <span onClick={() => setError(null)} style={{marginLeft:"auto",cursor:"pointer",
            fontWeight:700,color:"#b45309"}}>✕</span>
        </div>
      )}
      {savedOk && (
        <div style={{background:"#f0fdf4",borderBottom:"1px solid #bbf7d0",
          padding:"7px 18px",fontSize:12.5,color:"#166534",flexShrink:0}}>
          ✓ Coordonate salvate pentru <b>{savedOk}</b>
        </div>
      )}
      {saving && (
        <div style={{background:"#eff6ff",borderBottom:"1px solid #bfdbfe",
          padding:"7px 18px",fontSize:12.5,color:"#1d4ed8",flexShrink:0}}>
          ⏳ Se salvează coordonatele pentru <b>{saving}</b>...
        </div>
      )}

      {/* ── CORP ──────────────────────────────────────────────────────────── */}
      <div style={{flex:1,display:"flex",overflow:"hidden"}}>

        {/* ── SIDEBAR ─────────────────────────────────────────────────────── */}
        <div style={{
          width:272,background:"#fff",borderRight:"1px solid #e2e8f0",
          display:"flex",flexDirection:"column",overflow:"hidden",flexShrink:0,
        }}>
          {/* Filtre */}
          <div style={{padding:"11px 12px",borderBottom:"1px solid #f1f5f9"}}>
            <input value={searchQ} onChange={e => setSearchQ(e.target.value)}
              placeholder="Caută contor, drum, localitate..."
              style={{width:"100%",padding:"6px 10px",border:"1.5px solid #e2e8f0",
                borderRadius:7,fontSize:12,outline:"none",boxSizing:"border-box",
                background:"#f8fafc"}}/>

            {/* Status btns */}
            <div style={{marginTop:9,display:"flex",flexWrap:"wrap",gap:4}}>
              <button onClick={() => setFilterSt("ALL")} style={{
                padding:"3px 8px",borderRadius:11,fontSize:10.5,fontWeight:600,
                cursor:"pointer",
                background: filterSt==="ALL" ? "#0f172a" : "transparent",
                color: filterSt==="ALL" ? "#fff" : "#64748b",
                border:"1.5px solid #cbd5e1",
              }}>Toate</button>
              {Object.entries(STATUS_CFG).map(([k,v]) => (
                <button key={k} onClick={() => setFilterSt(f => f===k?"ALL":k)} style={{
                  padding:"3px 8px",borderRadius:11,fontSize:10.5,fontWeight:600,
                  cursor:"pointer",
                  background: filterSt===k ? v.fill : "transparent",
                  color: filterSt===k ? "#fff" : v.color,
                  border:`1.5px solid ${v.ring}`,transition:"all .15s",
                }}>{v.label}</button>
              ))}
            </div>

            {/* Tip filter */}
            {tipuri.length > 1 && (
              <div style={{marginTop:7,display:"flex",gap:4}}>
                {["ALL",...tipuri].map(t => (
                  <button key={t} onClick={() => setFilterTip(f => f===t?"ALL":t)} style={{
                    padding:"3px 9px",borderRadius:11,fontSize:10.5,fontWeight:600,
                    cursor:"pointer",
                    background: filterTip===t ? "#3b82f6" : "transparent",
                    color: filterTip===t ? "#fff" : "#3b82f6",
                    border:"1.5px solid #93c5fd",
                  }}>{t==="ALL"?"Toate":t}</button>
                ))}
              </div>
            )}

            <div style={{marginTop:7,fontSize:11,color:"#94a3b8"}}>
              {filtered.length} din {contoare.length} contoare
            </div>
          </div>

          {/* Lista */}
          <div style={{flex:1,overflowY:"auto"}}>
            {loading ? (
              <div style={{padding:28,textAlign:"center",color:"#94a3b8",fontSize:13}}>
                <div style={{fontSize:26,marginBottom:8}}>⏳</div>Se încarcă...
              </div>
            ) : filtered.length === 0 ? (
              <div style={{padding:18,textAlign:"center",color:"#94a3b8",fontSize:13}}>
                Niciun contor corespunde filtrelor.
              </div>
            ) : filtered.map(c => {
              const s    = STATUS_CFG[c.status] || STATUS_CFG.NELOCALIZAT;
              const isSel= selected === c.contor;
              const isSav= saving === c.contor;
              return (
                <div key={c.contor} onClick={() => focusContor(c)} style={{
                  padding:"9px 12px",borderBottom:"1px solid #f8fafc",
                  cursor: c.lat ? "pointer" : "default",
                  background: isSel ? "#eff6ff" : isSav ? "#f0fdf4" : "transparent",
                  borderLeft:`3px solid ${isSel ? "#3b82f6" : "transparent"}`,
                  transition:"background .1s",
                }}>
                  <div style={{display:"flex",alignItems:"center",gap:7,marginBottom:2}}>
                    <span style={{
                      width:8,height:8,borderRadius:"50%",
                      background:s.fill,flexShrink:0,
                      boxShadow:`0 0 4px ${s.fill}99`,
                    }}/>
                    <span style={{fontWeight:700,fontSize:12.5,color:"#0f172a"}}>{c.contor}</span>
                    <span style={{marginLeft:"auto",fontSize:9.5,fontWeight:600,
                      color:"#64748b",background:"#f1f5f9",padding:"1px 5px",borderRadius:7}}>
                      {c.tip || "—"}
                    </span>
                  </div>
                  <div style={{fontSize:11,color:"#475569",paddingLeft:15}}>
                    {c.drum} · {c.pozitie_km || "km?"} · {c.localitate || "—"}
                  </div>
                  <div style={{fontSize:11,fontWeight:600,color:s.fill,paddingLeft:15,marginTop:1}}>
                    MZL: {c.mzl != null ? fmt(c.mzl)+" veh/zi" : "—"}
                    {c.mzl_sursa === "manual" && (
                      <span style={{fontSize:9.5,background:"#dbeafe",color:"#1d4ed8",
                        padding:"1px 5px",borderRadius:7,marginLeft:4,fontWeight:600}}>M</span>
                    )}
                    {c.luna_ref && (
                      <span style={{color:"#94a3b8",fontWeight:400}}> · {c.luna_ref}</span>
                    )}
                  </div>
                  {!c.lat && (
                    <div style={{fontSize:10,color:"#7c3aed",paddingLeft:15,marginTop:1}}>
                      📍 Trage markerul din centrul hărții
                    </div>
                  )}
                </div>
              );
            })}
          </div>

          {/* Legendă */}
          <div style={{padding:"9px 12px",borderTop:"1px solid #e2e8f0",background:"#fafafa"}}>
            <div style={{fontSize:10,fontWeight:700,color:"#64748b",
              marginBottom:5,letterSpacing:0.5}}>LEGENDĂ</div>
            {Object.entries(STATUS_CFG).map(([k,v]) => (
              <div key={k} style={{display:"flex",alignItems:"center",gap:6,marginBottom:3}}>
                <span style={{width:8,height:8,borderRadius:"50%",
                  background:v.fill,flexShrink:0}}/>
                <span style={{fontSize:11,color:"#475569"}}>{v.label}</span>
              </div>
            ))}
            <div style={{marginTop:5,fontSize:10,color:"#94a3b8",lineHeight:1.4}}>
              Clasa 15 ≥ 10% → Totalizator<br/>
              Contoare fără coordonate: trage markerul gri pe poziție
            </div>
          </div>
        </div>

        {/* ── HARTĂ ─────────────────────────────────────────────────────────── */}
        <div style={{flex:1,position:"relative"}}>
          <div ref={mapRef} style={{width:"100%",height:"100%"}}/>

          {/* Badge nelocalizate */}
          {contoare.filter(c => !c.lat || !c.lng).length > 0 && (
            <div style={{
              position:"absolute",top:12,left:12,zIndex:1000,
              background:"rgba(124,58,237,0.12)",backdropFilter:"blur(8px)",
              border:"1px solid #7c3aed",borderRadius:8,
              padding:"6px 12px",fontSize:11.5,color:"#7c3aed",fontWeight:600,
            }}>
              {contoare.filter(c => !c.lat||!c.lng).length} contor(e) nelocalizate —
              trage markerul <span style={{opacity:.7}}>⊕</span> pe poziție
            </div>
          )}

          {!leafletReady && (
            <div style={{position:"absolute",inset:0,display:"flex",
              alignItems:"center",justifyContent:"center",
              background:"#f1f5f9",zIndex:999,fontSize:14,color:"#64748b"}}>
              Se încarcă harta...
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
