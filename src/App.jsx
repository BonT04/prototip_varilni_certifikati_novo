import { useState, useCallback, useRef } from "react";
import * as XLSX from "xlsx";

const WELDING_PROCESSES = [
  "WELDING POWER SOURCE - GMAW",
  "WELDING POWER SOURCE - TIG",
  "WELDING POWER SOURCE - SAW",
  "WELDING POWER SOURCE - SW",
  "WELDING POWER SOURCE -  MMA",
  "WELDING POWER SOURCE -  PLASMA",
  "RESISTANCE SPOT WELDING -  RSW",
  "VIR VARILNEGA TOKA - UPOROVNI STROJ",
];

const MANUFACTURERS = [
  "CLOOS","DAIHEN VARSTROJ","ESAB","ESS","EWM","FRONIUS","INVERTEC","ISKRA",
  "JÄCKLE","KEMPPI","LINCOLN","LORCH","MESSER GRIESHEIM","MIGATRONIC",
  "OERLIKON","RADE KONČAR","REHM","VARSTROJ","WEGA","WTL",
];

const INITIAL_HEADER = {
  datum: "2026-02-27",
  zacetna_st: 2559,
  narocnik: "PALFINGER proizvodnja d.o.o.",
  naslov: "Špelina ul. 22, 2000 Maribor",
  narocilo: "5101-4511142228",
  DN: 4161,
  porocilo_st: "1366/T-26",
  kraj_datum: "Maribor, 27.02.2026",
  temperatura: 20,
  vlaga: 43,
  datum_nalepke: "2026-02-27",
};

const INITIAL_BLOCKS = [
  {"seq":1,"tip":"TRANSTIG 2600","postopek":"WELDING POWER SOURCE - TIG","tov_st":"17301936","inv_st":"12206","proizvajalec":"FRONIUS","presek":4,"padec":0.5,"vhodni_25":200,"varilni_25":200,"vhodni_5":200,"U0x":83,"U0s":78,"U21x":10.6,"U21s":10.3,"U22x":17.5,"U22x2":30,"U22s":17.1,"U22s2":29.5,"I2x1":15,"I2s1":14.6,"I2x2":140,"I2s2":140.3,"I2x3":270,"I2s3":269.8,"Imax":260,"start_row":12},
  {"seq":2,"tip":"TRANSTIG 3000","postopek":"WELDING POWER SOURCE - TIG","tov_st":"11420011","inv_st":"11934","proizvajalec":"FRONIUS","presek":4,"padec":0.3,"vhodni_25":200,"varilni_25":200,"vhodni_5":200,"U0x":55,"U0s":52,"U21x":11.3,"U21s":11,"U22x":18.5,"U22x2":25.8,"U22s":18.1,"U22s2":25.2,"I2x1":16,"I2s1":16.3,"I2x2":151,"I2s2":149.5,"I2x3":300,"I2s3":298,"Imax":300,"start_row":38},
  {"seq":3,"tip":"TRANSTIG 2200","postopek":"WELDING POWER SOURCE - TIG","tov_st":"21350181","inv_st":"12343","proizvajalec":"FRONIUS","presek":4,"padec":0.11,"vhodni_25":200,"varilni_25":200,"vhodni_5":200,"U0x":89,"U0s":50,"U21x":10.8,"U21s":10.5,"U22x":17.1,"U22x2":23.2,"U22s":16.8,"U22s2":22.7,"I2x1":16,"I2s1":15.9,"I2x2":121,"I2s2":120.3,"I2x3":250,"I2s3":248.9,"Imax":250,"start_row":64},
  {"seq":4,"tip":"TRANSTIG 3000","postopek":"WELDING POWER SOURCE - TIG","tov_st":"11420012","inv_st":"11941","proizvajalec":"FRONIUS","presek":4,"padec":0.3,"vhodni_25":200,"varilni_25":200,"vhodni_5":200,"U0x":55,"U0s":62.8,"U21x":10.7,"U21s":10.4,"U22x":18.6,"U22x2":23.7,"U22s":18.3,"U22s2":23.2,"I2x1":15,"I2s1":15.5,"I2x2":148,"I2s2":147,"I2x3":300,"I2s3":298.3,"Imax":300,"start_row":90},
  {"seq":5,"tip":"TRANSTIG 3000","postopek":"WELDING POWER SOURCE - TIG","tov_st":"11420014","inv_st":"11944","proizvajalec":"FRONIUS","presek":4,"padec":0.35,"vhodni_25":200,"varilni_25":200,"vhodni_5":200,"U0x":56,"U0s":52.6,"U21x":11,"U21s":10.8,"U22x":18.3,"U22x2":29.5,"U22s":18,"U22s2":29.1,"I2x1":16,"I2s1":16.6,"I2x2":151,"I2s2":151,"I2x3":300,"I2s3":300,"Imax":300,"start_row":116},
  {"seq":6,"tip":"TRANSTIG 3000","postopek":"WELDING POWER SOURCE - TIG","tov_st":"11420010","inv_st":"11943","proizvajalec":"FRONIUS","presek":4,"padec":0.36,"vhodni_25":200,"varilni_25":200,"vhodni_5":200,"U0x":56,"U0s":52,"U21x":10.5,"U21s":10.3,"U22x":16,"U22x2":19.2,"U22s":15.6,"U22s2":18.9,"I2x1":26,"I2s1":25.7,"I2x2":146,"I2s2":146.5,"I2x3":300,"I2s3":301,"Imax":300,"start_row":142},
  {"seq":7,"tip":"TRANSTIG 3000","postopek":"WELDING POWER SOURCE - TIG","tov_st":"14451950","inv_st":"12090","proizvajalec":"FRONIUS","presek":4,"padec":0.35,"vhodni_25":200,"varilni_25":200,"vhodni_5":200,"U0x":56,"U0s":78,"U21x":10.9,"U21s":10.7,"U22x":17.6,"U22x2":24,"U22s":17.3,"U22s2":23.6,"I2x1":16,"I2s1":16,"I2x2":157,"I2s2":156.6,"I2x3":300,"I2s3":299.6,"Imax":300,"start_row":168},
  {"seq":8,"tip":"TRANTIG 3000","postopek":"WELDING POWER SOURCE - TIG","tov_st":"11420016","inv_st":"11942","proizvajalec":"FRONIUS","presek":4,"padec":0.35,"vhodni_25":200,"varilni_25":200,"vhodni_5":200,"U0x":56,"U0s":52.5,"U21x":11,"U21s":10.8,"U22x":18.1,"U22x2":24.5,"U22s":17.8,"U22s2":24,"I2x1":16,"I2s1":16.4,"I2x2":157,"I2s2":156,"I2x3":300,"I2s3":299,"Imax":300,"start_row":194},
  {"seq":9,"tip":"TRANTIG 2500","postopek":"WELDING POWER SOURCE - TIG","tov_st":"24393094","inv_st":"12419","proizvajalec":"FRONIUS","presek":4,"padec":0.42,"vhodni_25":200,"varilni_25":200,"vhodni_5":200,"U0x":89,"U0s":50,"U21x":10.1,"U21s":10.9,"U22x":18.9,"U22x2":24.4,"U22s":18.4,"U22s2":23.8,"I2x1":15,"I2s1":14.9,"I2x2":150,"I2s2":149,"I2x3":250,"I2s3":248,"Imax":250,"start_row":220},
  {"seq":10,"tip":"TRANSPULS SYNERGIC 3200","postopek":"WELDING POWER SOURCE - GMAW","tov_st":"15512541","inv_st":"12162","proizvajalec":"FRONIUS","presek":4,"padec":0.3,"vhodni_25":200,"varilni_25":200,"vhodni_5":200,"U0x":78,"U0s":40,"U21x":15.8,"U21s":16,"U22x":21.6,"U22x2":28.4,"U22s":21.9,"U22s2":28.9,"I2x1":45,"I2s1":45.1,"I2x2":151,"I2s2":152,"I2x3":320,"I2s3":320,"Imax":320,"start_row":246},
  {"seq":11,"tip":"TRANSTIG 2600","postopek":"WELDING POWER SOURCE - TIG","tov_st":"17271457","inv_st":"12207","proizvajalec":"FRONIUS","presek":4,"padec":0.36,"vhodni_25":200,"varilni_25":200,"vhodni_5":200,"U0x":83,"U0s":78.2,"U21x":10.9,"U21s":10.9,"U22x":18,"U22x2":25.1,"U22s":17.8,"U22s2":24.7,"I2x1":16,"I2s1":16.1,"I2x2":132,"I2s2":132,"I2x3":260,"I2s3":259,"Imax":260,"start_row":272},
  {"seq":12,"tip":"TRANSTIG 3000","postopek":"WELDING POWER SOURCE - TIG","tov_st":"19153484","inv_st":"12276","proizvajalec":"FRONIUS","presek":4,"padec":0.32,"vhodni_25":200,"varilni_25":200,"vhodni_5":200,"U0x":85,"U0s":50,"U21x":11.1,"U21s":11,"U22x":18.6,"U22x2":23.6,"U22s":18.4,"U22s2":23.1,"I2x1":17,"I2s1":16.8,"I2x2":150,"I2s2":149.4,"I2x3":300,"I2s3":298.9,"Imax":300,"start_row":298},
  {"seq":13,"tip":"TRANSTIG 2500","postopek":"WELDING POWER SOURCE - TIG","tov_st":"22162448","inv_st":"12342","proizvajalec":"FRONIUS","presek":4,"padec":0.33,"vhodni_25":200,"varilni_25":200,"vhodni_5":200,"U0x":89,"U0s":50,"U21x":9.9,"U21s":9.7,"U22x":19.2,"U22x2":27.2,"U22s":18.6,"U22s2":26.5,"I2x1":15,"I2s1":14.9,"I2x2":129,"I2s2":128,"I2x3":250,"I2s3":248.6,"Imax":250,"start_row":324},
  {"seq":14,"tip":"MAGICWAVE 3000","postopek":"WELDING POWER SOURCE - TIG","tov_st":"19153486","inv_st":"12275","proizvajalec":"FRONIUS","presek":4,"padec":0.6,"vhodni_25":200,"varilni_25":200,"vhodni_5":200,"U0x":89,"U0s":50,"U21x":9.9,"U21s":9.6,"U22x":17.4,"U22x2":25.6,"U22s":17,"U22s2":24.8,"I2x1":15,"I2s1":14.8,"I2x2":150,"I2s2":148,"I2x3":300,"I2s3":297,"Imax":300,"start_row":350},
  {"seq":15,"tip":"TRANSTIG 3000","postopek":"WELDING POWER SOURCE - TIG","tov_st":"11420017","inv_st":"12207","proizvajalec":"FRONIUS","presek":4,"padec":0.61,"vhodni_25":200,"varilni_25":200,"vhodni_5":200,"U0x":56,"U0s":52.7,"U21x":10.9,"U21s":10.6,"U22x":25.1,"U22x2":27.8,"U22s":25,"U22s2":27.2,"I2x1":17,"I2s1":17,"I2x2":158,"I2s2":157,"I2x3":300,"I2s3":300,"Imax":300,"start_row":376},
  {"seq":16,"tip":"TRANSTIG 3000","postopek":"WELDING POWER SOURCE - TIG","tov_st":"11040090","inv_st":"11893","proizvajalec":"FRONIUS","presek":4,"padec":0.45,"vhodni_25":200,"varilni_25":200,"vhodni_5":200,"U0x":56,"U0s":52.5,"U21x":10.9,"U21s":10.6,"U22x":18.5,"U22x2":23.9,"U22s":18.2,"U22s2":23.4,"I2x1":16,"I2s1":16.3,"I2x2":152,"I2s2":153,"I2x3":300,"I2s3":302,"Imax":300,"start_row":402},
  {"seq":17,"tip":"TRANSTIG 210","postopek":"WELDING POWER SOURCE - TIG","tov_st":"31045678","inv_st":"/","proizvajalec":"FRONIUS","presek":2.5,"padec":"N/A","vhodni_25":200,"varilni_25":200,"vhodni_5":200,"U0x":101,"U0s":86,"U21x":13.1,"U21s":12.8,"U22x":18.6,"U22x2":23.7,"U22s":18,"U22s2":22.8,"I2x1":21,"I2s1":20.5,"I2x2":90,"I2s2":91,"I2x3":180,"I2s3":181,"Imax":210,"start_row":428},
  {"seq":18,"tip":"TRANSTIG 3000","postopek":"WELDING POWER SOURCE - TIG","tov_st":"14021649","inv_st":"11939","proizvajalec":"FRONIUS","presek":4,"padec":0.43,"vhodni_25":200,"varilni_25":200,"vhodni_5":200,"U0x":83,"U0s":79.3,"U21x":10.5,"U21s":10.2,"U22x":18.9,"U22x2":27,"U22s":18.6,"U22s2":26.5,"I2x1":16,"I2s1":16.5,"I2x2":149,"I2s2":149,"I2x3":300,"I2s3":300,"Imax":300,"start_row":454},
  {"seq":19,"tip":"TRANSTIG 3000","postopek":"WELDING POWER SOURCE - TIG","tov_st":"11420015","inv_st":"11938","proizvajalec":"FRONIUS","presek":4,"padec":0.41,"vhodni_25":200,"varilni_25":200,"vhodni_5":200,"U0x":56,"U0s":53,"U21x":"/","U21s":null,"U22x":null,"U22x2":null,"U22s":null,"U22s2":null,"I2x1":16,"I2s1":16.2,"I2x2":153,"I2s2":152.5,"I2x3":300,"I2s3":299.8,"Imax":300,"start_row":480},
  {"seq":20,"tip":"TRANSTIG 2500","postopek":"WELDING POWER SOURCE - TIG","tov_st":"21315239","inv_st":"12309","proizvajalec":"FRONIUS","presek":4,"padec":0.5,"vhodni_25":200,"varilni_25":200,"vhodni_5":200,"U0x":89,"U0s":50,"U21x":11.2,"U21s":11,"U22x":18.3,"U22x2":25.6,"U22s":17.9,"U22s2":25,"I2x1":16,"I2s1":16,"I2x2":130,"I2s2":129,"I2x3":250,"I2s3":250,"Imax":250,"start_row":506},
];

const n = (v) => (v === null || v === undefined || v === "" ? "" : v);
const asNum = (v) => {
  if (v === null || v === undefined || v === "" || v === "/" || v === "N/A") return v;
  const parsed = parseFloat(v);
  return isNaN(parsed) ? v : parsed;
};

function FieldInput({ label, value, onChange, small, type = "text" }) {
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 2 }}>
      <label style={{ fontSize: 10, color: "#8fa3b1", fontFamily: "'Courier New', monospace", textTransform: "uppercase", letterSpacing: 1 }}>{label}</label>
      <input
        type={type}
        value={n(value)}
        onChange={e => onChange(e.target.value)}
        style={{
          background: "#1a2332",
          border: "1px solid #2d4059",
          borderRadius: 4,
          color: "#e8f4f8",
          padding: small ? "4px 6px" : "6px 8px",
          fontSize: small ? 13 : 14,
          fontFamily: "'Courier New', monospace",
          width: "100%",
          boxSizing: "border-box",
          outline: "none",
          transition: "border-color 0.15s",
        }}
        onFocus={e => e.target.style.borderColor = "#00d4ff"}
        onBlur={e => e.target.style.borderColor = "#2d4059"}
      />
    </div>
  );
}

function SelectInput({ label, value, onChange, options }) {
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 2 }}>
      <label style={{ fontSize: 10, color: "#8fa3b1", fontFamily: "'Courier New', monospace", textTransform: "uppercase", letterSpacing: 1 }}>{label}</label>
      <select
        value={value || ""}
        onChange={e => onChange(e.target.value)}
        style={{
          background: "#1a2332",
          border: "1px solid #2d4059",
          borderRadius: 4,
          color: "#e8f4f8",
          padding: "6px 8px",
          fontSize: 13,
          fontFamily: "'Courier New', monospace",
          width: "100%",
          boxSizing: "border-box",
          outline: "none",
        }}
      >
        <option value="">— izberi —</option>
        {options.map(o => <option key={o} value={o}>{o}</option>)}
      </select>
    </div>
  );
}

function MeasRow({ label, xVal, sVal, onXChange, onSChange }) {
  return (
    <div style={{ display: "grid", gridTemplateColumns: "120px 1fr 1fr", gap: 6, alignItems: "end" }}>
      <span style={{ fontSize: 11, color: "#8fa3b1", fontFamily: "'Courier New', monospace", paddingBottom: 4 }}>{label}</span>
      <FieldInput label="x (izvor)" value={xVal} onChange={onXChange} small />
      <FieldInput label="s (merilnik)" value={sVal} onChange={onSChange} small />
    </div>
  );
}

function DeviceCard({ block, idx, onChange, isActive, onSelect }) {
  const upd = (key) => (val) => onChange(idx, key, val);
  return (
    <div
      onClick={onSelect}
      style={{
        background: isActive ? "#0d1f33" : "#0a1628",
        border: isActive ? "1px solid #00d4ff" : "1px solid #1e3050",
        borderRadius: 8,
        padding: "14px 16px",
        cursor: "pointer",
        transition: "all 0.15s",
        boxShadow: isActive ? "0 0 12px rgba(0,212,255,0.15)" : "none",
      }}
    >
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: isActive ? 16 : 0 }}>
        <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
          <span style={{
            background: isActive ? "#00d4ff" : "#1e3050",
            color: isActive ? "#000" : "#8fa3b1",
            borderRadius: 4,
            padding: "2px 8px",
            fontSize: 12,
            fontWeight: 700,
            fontFamily: "'Courier New', monospace",
          }}>#{block.seq}</span>
          <span style={{ color: isActive ? "#e8f4f8" : "#7a9ab5", fontSize: 14, fontFamily: "'Courier New', monospace" }}>
            {block.tip || <span style={{ color: "#4a6a8a", fontStyle: "italic" }}>brez tipa</span>}
          </span>
          <span style={{ color: "#4a6a8a", fontSize: 11 }}>{block.tov_st}</span>
        </div>
        <span style={{ color: isActive ? "#00d4ff" : "#4a6a8a", fontSize: 12 }}>{isActive ? "▲ zapri" : "▼ uredi"}</span>
      </div>

      {isActive && (
        <div onClick={e => e.stopPropagation()}>
          {/* Identity */}
          <div style={{ background: "#0f1f35", borderRadius: 6, padding: 12, marginBottom: 10 }}>
            <div style={{ color: "#00d4ff", fontSize: 11, fontFamily: "'Courier New', monospace", marginBottom: 8, textTransform: "uppercase", letterSpacing: 2 }}>■ Identifikacija</div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
              <FieldInput label="Tip varilnega izvora" value={block.tip} onChange={upd("tip")} />
              <SelectInput label="Postopek" value={block.postopek} onChange={upd("postopek")} options={WELDING_PROCESSES} />
              <FieldInput label="Tovarniška št." value={block.tov_st} onChange={upd("tov_st")} />
              <FieldInput label="Inventarna št." value={block.inv_st} onChange={upd("inv_st")} />
              <SelectInput label="Proizvajalec" value={block.proizvajalec} onChange={upd("proizvajalec")} options={MANUFACTURERS} />
              <FieldInput label="Presek kabla [mm²]" value={block.presek} onChange={upd("presek")} />
            </div>
          </div>

          {/* Section 1.1 */}
          <div style={{ background: "#0f1f35", borderRadius: 6, padding: 12, marginBottom: 10 }}>
            <div style={{ color: "#00d4ff", fontSize: 11, fontFamily: "'Courier New', monospace", marginBottom: 8, textTransform: "uppercase", letterSpacing: 2 }}>■ 1.1 Padec napetosti</div>
            <FieldInput label="Padec napetosti Uc [V]" value={block.padec} onChange={upd("padec")} />
          </div>

          {/* Section 1.2 */}
          <div style={{ background: "#0f1f35", borderRadius: 6, padding: 12, marginBottom: 10 }}>
            <div style={{ color: "#00d4ff", fontSize: 11, fontFamily: "'Courier New', monospace", marginBottom: 8, textTransform: "uppercase", letterSpacing: 2 }}>■ 1.2 Izolacijska upornost [MΩ]</div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 8 }}>
              <FieldInput label="Vhodni ≥ 2,5 MΩ" value={block.vhodni_25} onChange={upd("vhodni_25")} small />
              <FieldInput label="Varilni ≥ 2,5 MΩ" value={block.varilni_25} onChange={upd("varilni_25")} small />
              <FieldInput label="Vhodni ≥ 5 MΩ" value={block.vhodni_5} onChange={upd("vhodni_5")} small />
            </div>
          </div>

          {/* Section 1.3 */}
          <div style={{ background: "#0f1f35", borderRadius: 6, padding: 12, marginBottom: 10 }}>
            <div style={{ color: "#00d4ff", fontSize: 11, fontFamily: "'Courier New', monospace", marginBottom: 8, textTransform: "uppercase", letterSpacing: 2 }}>■ 1.3 Napetost prostega teka U0</div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
              <FieldInput label="U0x [V]" value={block.U0x} onChange={upd("U0x")} small />
              <FieldInput label="U0s [V]" value={block.U0s} onChange={upd("U0s")} small />
            </div>
          </div>

          {/* Section 1.4 */}
          <div style={{ background: "#0f1f35", borderRadius: 6, padding: 12, marginBottom: 10 }}>
            <div style={{ color: "#00d4ff", fontSize: 11, fontFamily: "'Courier New', monospace", marginBottom: 8, textTransform: "uppercase", letterSpacing: 2 }}>■ 1.4 Napetost U2</div>
            <MeasRow label="U21 [V]" xVal={block.U21x} sVal={block.U21s} onXChange={upd("U21x")} onSChange={upd("U21s")} />
            <div style={{ height: 6 }} />
            <div style={{ display: "grid", gridTemplateColumns: "120px 1fr 1fr 1fr 1fr", gap: 6, alignItems: "end" }}>
              <span style={{ fontSize: 11, color: "#8fa3b1", fontFamily: "'Courier New', monospace", paddingBottom: 4 }}>U22 [V]</span>
              <FieldInput label="x1 (izvor)" value={block.U22x} onChange={upd("U22x")} small />
              <FieldInput label="x2 (izvor)" value={block.U22x2} onChange={upd("U22x2")} small />
              <FieldInput label="s1 (merilnik)" value={block.U22s} onChange={upd("U22s")} small />
              <FieldInput label="s2 (merilnik)" value={block.U22s2} onChange={upd("U22s2")} small />
            </div>
          </div>

          {/* Section 1.5 */}
          <div style={{ background: "#0f1f35", borderRadius: 6, padding: 12 }}>
            <div style={{ color: "#00d4ff", fontSize: 11, fontFamily: "'Courier New', monospace", marginBottom: 8, textTransform: "uppercase", letterSpacing: 2 }}>■ 1.5 Tok I2</div>
            <MeasRow label="I2-1 [A]" xVal={block.I2x1} sVal={block.I2s1} onXChange={upd("I2x1")} onSChange={upd("I2s1")} />
            <div style={{ height: 6 }} />
            <MeasRow label="I2-2 [A]" xVal={block.I2x2} sVal={block.I2s2} onXChange={upd("I2x2")} onSChange={upd("I2s2")} />
            <div style={{ height: 6 }} />
            <MeasRow label="I2-3 [A]" xVal={block.I2x3} sVal={block.I2s3} onXChange={upd("I2x3")} onSChange={upd("I2s3")} />
            <div style={{ height: 8 }} />
            <FieldInput label="Imax [A]" value={block.Imax} onChange={upd("Imax")} small />
          </div>
        </div>
      )}
    </div>
  );
}

export default function App() {
  const [header, setHeader] = useState(INITIAL_HEADER);
  const [blocks, setBlocks] = useState(INITIAL_BLOCKS.map(b => ({ ...b })));
  const [activeIdx, setActiveIdx] = useState(null);
  const [tab, setTab] = useState("devices"); // "header" | "devices"
  const [saved, setSaved] = useState(false);

  const updHeader = (key) => (val) => setHeader(h => ({ ...h, [key]: val }));
  const updBlock = useCallback((idx, key, val) => {
    setBlocks(bs => bs.map((b, i) => i === idx ? { ...b, [key]: val } : b));
  }, []);

  const [search, setSearch] = useState("");
  const filteredBlocks = blocks
    .map((b, i) => ({ b, i }))
    .filter(({ b }) => {
      if (!search) return true;
      const q = search.toLowerCase();
      return (
        String(b.tov_st || "").toLowerCase().includes(q) ||
        String(b.inv_st || "").toLowerCase().includes(q) ||
        String(b.tip || "").toLowerCase().includes(q) ||
        String(b.proizvajalec || "").toLowerCase().includes(q) ||
        String(b.seq || "").includes(q)
      );
    });

  const [showInstructions, setShowInstructions] = useState(false);
  const [showJSON, setShowJSON] = useState(false);
  const [copied, setCopied] = useState(false);

  const getJSON = () => JSON.stringify({ header, blocks }, null, 2);

  const copyJSON = () => {
    navigator.clipboard.writeText(getJSON()).then(() => {
      setCopied(true);
      setTimeout(() => setCopied(false), 3000);
    });
  };

  const [importMsg, setImportMsg] = useState(null);
  const fileInputRef = useRef();

  const handleImportXLSX = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setImportMsg({ type: "loading", text: "Berem datoteko..." });
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const wb = XLSX.read(ev.target.result, { type: "array", cellDates: true });
        const ws = wb.Sheets["Osnova"];
        if (!ws) throw new Error("List 'Osnova' ni najden v datoteki.");
        const v = (r, c) => { const cell = ws[XLSX.utils.encode_cell({r: r-1, c: c-1})]; return cell ? cell.v : null; };
        const s = (r, c) => { const cell = ws[XLSX.utils.encode_cell({r: r-1, c: c-1})]; return cell ? String(cell.v ?? "") : ""; };

        // Parse header
        const newHeader = {
          datum: s(1, 2),
          zacetna_st: v(1, 5),
          narocnik: s(2, 2),
          naslov: s(3, 2),
          narocilo: s(4, 2),
          DN: v(5, 2),
          porocilo_st: s(6, 2),
          kraj_datum: s(7, 2),
          temperatura: v(8, 2),
          vlaga: v(9, 2),
          datum_nalepke: s(10, 2),
        };

        // Find blocks by scanning for "Zaporedna št." in col A
        const newBlocks = [];
        const range = XLSX.utils.decode_range(ws["!ref"]);
        for (let row = 1; row <= range.e.r + 1; row++) {
          const cell = ws[XLSX.utils.encode_cell({r: row-1, c: 0})];
          if (cell && cell.v === "Zaporedna št.") {
            const sr = row;
            newBlocks.push({
              seq:          v(sr,   2),
              tip:          s(sr+2, 2),
              postopek:     s(sr+3, 2),
              tov_st:       s(sr+4, 2),
              inv_st:       s(sr+5, 2),
              proizvajalec: s(sr+6, 2),
              presek:       v(sr+7, 2),
              padec:        v(sr+8, 2),
              vhodni_25:    v(sr+9, 2),
              varilni_25:   v(sr+10, 2),
              vhodni_5:     v(sr+11, 2),
              U0x:          v(sr+12, 2),
              U0s:          v(sr+13, 2),
              U21x:         v(sr+14, 2),
              U21s:         v(sr+15, 2),
              U22x:         v(sr+16, 2),
              U22x2:        v(sr+16, 3),
              U22s:         v(sr+17, 2),
              U22s2:        v(sr+17, 3),
              I2x1:         v(sr+18, 2),
              I2s1:         v(sr+19, 2),
              I2x2:         v(sr+20, 2),
              I2s2:         v(sr+21, 2),
              I2x3:         v(sr+22, 2),
              I2s3:         v(sr+23, 2),
              Imax:         v(sr+24, 2),
              start_row:    sr,
            });
          }
        }

        if (newBlocks.length === 0) throw new Error("Ni najdenih blokov naprav.");
        setHeader(newHeader);
        setBlocks(newBlocks);
        setActiveIdx(null);
        setSearch("");
        setImportMsg({ type: "ok", text: `✓ Uvoženo ${newBlocks.length} naprav iz "${file.name}"` });
        setTimeout(() => setImportMsg(null), 5000);
      } catch (err) {
        setImportMsg({ type: "err", text: "Napaka: " + err.message });
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = "";
  };

  return (
    <div style={{
      background: '#060f1e',
      minHeight: '100vh',
      color: '#e8f4f8',
      fontFamily: "'Courier New', monospace",
      padding: '0 0 40px 0',
    }}>
      <div style={{
        background: 'linear-gradient(135deg, #0a1e35, #0d2a45)',
        borderBottom: '1px solid #1e3a5f',
        padding: '16px 24px',
        display: 'flex',
        justifyContent: 'space-between',
        alignItems: 'center',
        position: 'sticky',
        top: 0,
        zIndex: 100,
      }}>
        <div>
          <div style={{ color: '#00d4ff', fontSize: 11, letterSpacing: 4, textTransform: 'uppercase', marginBottom: 2 }}>Varilni izvori · Vnos podatkov</div>
          <div style={{ fontSize: 18, fontWeight: 700, color: '#e8f4f8' }}>PALFINGER — {header.porocilo_st}</div>
        </div>
        <div style={{ display: 'flex', gap: 8 }}>
          <button onClick={() => setShowInstructions(v => !v)} style={{
            background: 'transparent', color: '#8fa3b1', border: '1px solid #2d4059',
            borderRadius: 6, padding: '10px 16px', fontFamily: "'Courier New', monospace",
            fontSize: 12, cursor: 'pointer', letterSpacing: 1,
          }}>? NAVODILA</button>
          <input ref={fileInputRef} type="file" accept=".xlsx" style={{ display: 'none' }} onChange={handleImportXLSX} />
          <button onClick={() => fileInputRef.current.click()} style={{
            background: '#1a3a1a', color: '#7fffb2', border: '1px solid #2d6a2d',
            borderRadius: 6, padding: '10px 16px', fontFamily: "'Courier New', monospace",
            fontWeight: 700, fontSize: 13, cursor: 'pointer', letterSpacing: 1,
          }}>⬆ UVOZI XLSX</button>
          <button onClick={() => setShowJSON(v => !v)} style={{
            background: showJSON ? '#1a3a55' : '#00d4ff', color: showJSON ? '#00d4ff' : '#000', border: showJSON ? '1px solid #00d4ff' : 'none',
            borderRadius: 6, padding: '10px 20px', fontFamily: "'Courier New', monospace",
            fontWeight: 700, fontSize: 13, cursor: 'pointer', letterSpacing: 1,
          }}>{showJSON ? '✕ ZAPRI JSON' : '{ } PRIKAŽI JSON'}</button>
          <button onClick={copyJSON} style={{
            background: copied ? '#00a651' : '#2d6a2d', color: '#fff', border: 'none',
            borderRadius: 6, padding: '10px 20px', fontFamily: "'Courier New', monospace",
            fontWeight: 700, fontSize: 13, cursor: 'pointer', letterSpacing: 1,
          }}>{copied ? '✓ KOPIRANO' : '⧉ KOPIRAJ JSON'}</button>
        </div>
      </div>

      {importMsg && (
        <div style={{
          background: importMsg.type === "ok" ? "#0a2a0a" : importMsg.type === "err" ? "#2a0a0a" : "#0a1e35",
          border: `1px solid ${importMsg.type === "ok" ? "#2d6a2d" : importMsg.type === "err" ? "#6a2d2d" : "#1e4060"}`,
          borderRadius: 6, margin: '12px auto', padding: '10px 20px', maxWidth: 860,
          color: importMsg.type === "ok" ? "#7fffb2" : importMsg.type === "err" ? "#ff7f7f" : "#8fa3b1",
          fontSize: 13, fontFamily: "'Courier New', monospace",
        }}>{importMsg.text}</div>
      )}

      {showInstructions && (
        <div style={{
          background: '#0a1e35', border: '1px solid #1e4060', borderRadius: 8,
          margin: '16px auto', padding: 20, maxWidth: 860,
        }}>
          <div style={{ color: '#00d4ff', fontSize: 12, letterSpacing: 2, marginBottom: 12 }}>■ NAVODILA ZA IZVOZ V XLSX</div>
          <div style={{ color: '#a0c0d8', fontSize: 13, lineHeight: 1.8 }}>
            <p style={{ margin: '0 0 8px' }}>Ker mora izvoz ohraniti vse formule, oblikovanje in vse liste (1–35) iz originala, izvoz poteka v dveh korakih:</p>
            <p style={{ margin: '0 0 4px' }}><span style={{color:'#00d4ff'}}>1.</span> Kliknite <strong style={{color:'#e8f4f8'}}>⬇ IZVOZI JSON</strong> — shrani se datoteka <code style={{color:'#ffd700'}}>palfinger_podatki.json</code></p>
            <p style={{ margin: '0 0 4px' }}><span style={{color:'#00d4ff'}}>2.</span> Na računalniku zaženite Python skripto <code style={{color:'#ffd700'}}>palfinger_export.py</code>:</p>
            <pre style={{ background: '#060f1e', borderRadius: 6, padding: '10px 14px', color: '#7fffb2', fontSize: 12, margin: '8px 0 12px', overflowX: 'auto' }}>python3 palfinger_export.py palfinger_podatki.json palfinger_glavni.xlsx izhod.xlsx</pre>
            <p style={{ margin: 0, color: '#6a8a9a', fontSize: 12 }}>Skripta kopira originalno datoteko in samo posodobi podatkovne celice — formule, oblikovanje in vsi listi ostanejo nespremenjeni.</p>
          </div>
        </div>
      )}

      {showJSON && (
        <div style={{ margin: '16px auto', maxWidth: 860, padding: '0 20px' }}>
          <div style={{ background: '#0a1e35', border: '1px solid #1e4060', borderRadius: 8, overflow: 'hidden' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '10px 16px', borderBottom: '1px solid #1e4060' }}>
              <span style={{ color: '#00d4ff', fontSize: 11, letterSpacing: 2 }}>■ JSON PODATKI — kopirajte in shranite kot palfinger_podatki.json</span>
              <button onClick={copyJSON} style={{
                background: copied ? '#00a651' : '#2d6a2d', color: '#fff', border: 'none',
                borderRadius: 4, padding: '6px 14px', fontFamily: "'Courier New', monospace",
                fontSize: 12, cursor: 'pointer',
              }}>{copied ? '✓ KOPIRANO' : '⧉ KOPIRAJ'}</button>
            </div>
            <textarea
              readOnly
              value={getJSON()}
              onClick={e => e.target.select()}
              style={{
                width: '100%', height: 300, background: '#060f1e', color: '#7fffb2',
                border: 'none', padding: 16, fontFamily: "'Courier New', monospace",
                fontSize: 11, lineHeight: 1.5, resize: 'vertical', boxSizing: 'border-box', outline: 'none',
              }}
            />
          </div>
        </div>
      )}

      {/* Tabs */}
      <div style={{ display: "flex", gap: 0, borderBottom: "1px solid #1e3050", background: "#08131f" }}>
        {[["devices", `Naprave (${blocks.length})`], ["header", "Splošni podatki"]].map(([key, label]) => (
          <button key={key} onClick={() => setTab(key)} style={{
            background: "none",
            border: "none",
            borderBottom: tab === key ? "2px solid #00d4ff" : "2px solid transparent",
            color: tab === key ? "#00d4ff" : "#5a7a9a",
            padding: "12px 24px",
            fontFamily: "'Courier New', monospace",
            fontSize: 13,
            cursor: "pointer",
            letterSpacing: 1,
          }}>{label}</button>
        ))}
      </div>

      <div style={{ padding: "20px 20px", maxWidth: 900, margin: "0 auto" }}>
        {tab === "header" && (
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
            <FieldInput label="Datum dokumenta" value={header.datum} onChange={updHeader("datum")} />
            <FieldInput label="Začetna št. potrdila" value={header.zacetna_st} onChange={updHeader("zacetna_st")} />
            <FieldInput label="Naročnik" value={header.narocnik} onChange={updHeader("narocnik")} />
            <FieldInput label="Naslov" value={header.naslov} onChange={updHeader("naslov")} />
            <FieldInput label="Naročilo" value={header.narocilo} onChange={updHeader("narocilo")} />
            <FieldInput label="DN" value={header.DN} onChange={updHeader("DN")} />
            <FieldInput label="Poročilo št." value={header.porocilo_st} onChange={updHeader("porocilo_st")} />
            <FieldInput label="Kraj in datum pregleda" value={header.kraj_datum} onChange={updHeader("kraj_datum")} />
            <FieldInput label="Temperatura [°C]" value={header.temperatura} onChange={updHeader("temperatura")} />
            <FieldInput label="Vlaga [%]" value={header.vlaga} onChange={updHeader("vlaga")} />
            <FieldInput label="Datum nalepke" value={header.datum_nalepke} onChange={updHeader("datum_nalepke")} />
          </div>
        )}

        {tab === "devices" && (
          <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
            {/* Search bar */}
            <div style={{ position: "relative", marginBottom: 4 }}>
              <input
                type="text"
                placeholder="Iskanje po serijski št., inv. št., tipu ali proizvajalcu..."
                value={search}
                onChange={e => { setSearch(e.target.value); setActiveIdx(null); }}
                style={{
                  width: "100%", boxSizing: "border-box",
                  background: "#0f1f35", border: "1px solid #2d4059",
                  borderRadius: 6, color: "#e8f4f8", padding: "10px 40px 10px 14px",
                  fontFamily: "'Courier New', monospace", fontSize: 13, outline: "none",
                }}
                onFocus={e => e.target.style.borderColor = "#00d4ff"}
                onBlur={e => e.target.style.borderColor = "#2d4059"}
              />
              {search && (
                <button onClick={() => setSearch("")} style={{
                  position: "absolute", right: 10, top: "50%", transform: "translateY(-50%)",
                  background: "none", border: "none", color: "#8fa3b1", cursor: "pointer", fontSize: 16,
                }}>✕</button>
              )}
            </div>
            <div style={{ color: "#5a7a9a", fontSize: 11, marginBottom: 4 }}>
              {search
                ? `${filteredBlocks.length} rezultat${filteredBlocks.length === 1 ? "" : filteredBlocks.length < 5 ? "i" : "ov"} od ${blocks.length} naprav`
                : `Kliknite napravo za urejanje · ${blocks.length} naprav skupaj`}
            </div>
            {filteredBlocks.length === 0 && (
              <div style={{ color: "#4a6a8a", fontSize: 13, textAlign: "center", padding: 32 }}>
                Ni rezultatov za "{search}"
              </div>
            )}
            {filteredBlocks.map(({ b, i }) => (
              <DeviceCard
                key={i}
                block={b}
                idx={i}
                onChange={updBlock}
                isActive={activeIdx === i}
                onSelect={() => setActiveIdx(activeIdx === i ? null : i)}
              />
            ))}
          </div>
        )}
      </div>
    </div>
  );
}
