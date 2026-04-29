// ============================================================
// MBSE Interface Master v5  — Handle Hover Visibility
// React + ReactFlow  |  CodeSandbox ready
// package.json: "reactflow": "^11.11.4"
// ============================================================

import React, { useState, useCallback, useRef, useEffect, memo } from "react";
import ReactFlow, {
  ReactFlowProvider,
  addEdge,
  useNodesState,
  useEdgesState,
  Controls,
  Background,
  MiniMap,
  Handle,
  Position,
  useReactFlow,
  Panel,
  NodeResizer,
  EdgeLabelRenderer,
  ConnectionMode,
  getSmoothStepPath,
} from "reactflow";
import "reactflow/dist/style.css";

// ─────────────────────────────────────────────────────────────
// HANDLE HOVER CSS
// 평소엔 숨김 → 노드에 마우스 올리면 나타남 → 선택시 항상 표시
// ─────────────────────────────────────────────────────────────
const HANDLE_CSS = `
  .react-flow__handle {
    opacity: 0 !important;
    transition: opacity 0.15s ease, transform 0.15s ease !important;
  }
  .react-flow__node:hover .react-flow__handle {
    opacity: 1 !important;
  }
  .react-flow__node.selected .react-flow__handle {
    opacity: 1 !important;
  }
  .react-flow__handle:hover {
    opacity: 1 !important;
    transform: scale(1.4) !important;
  }
  .react-flow__handle.connecting {
    opacity: 1 !important;
  }
`;

// ─────────────────────────────────────────────────────────────
// CONSTANTS
// ─────────────────────────────────────────────────────────────
const AREA_TYPES = ["Plant","Function","Package","Item"];
const AREA_COLORS = {
  Plant:    { bg:"rgba(219,234,254,0.45)", border:"#93c5fd", label:"#1d4ed8" },
  Function: { bg:"rgba(220,252,231,0.45)", border:"#86efac", label:"#15803d" },
  Package:  { bg:"rgba(254,249,195,0.45)", border:"#fde047", label:"#a16207" },
  Item:     { bg:"rgba(243,232,255,0.45)", border:"#d8b4fe", label:"#7e22ce" },
};

const EQUIPMENT_LIST = [
  "Tank","Pump","Pond","Heat Exchanger","Filter","Hopper","Decanter",
  "Cooling Tower","Clarifier","Classifier","Feed Box","Chemical Dosing",
  "Scrubber","Bag Filter","Reactor","Feed Bin","Gas Duct","Steel Structure",
];
const CONNECTION_LIST = ["Piping","Duct","Brench"];
const CONVEYOR_LIST   = ["Conveyor"];
const INSTRUMENT_CATS = { Flow:[], Pressure:[], Temperature:[], Level:[] };
const INSTR_TYPES     = ["Transmitter","Switch","Gauge"];

const FLUID_PRIMARY = ["Water","Slurry Water","Gas","Process Gas","Steam","Chemical"];
const FLUID_SECONDARY = {
  Water:          ["FW","PW","WM","GW","SW","WW","MCWS","MCR","PWS","PWR","PCC"],
  "Slurry Water": ["PWR","PCC","SL","TW"],
  Gas:            ["NI","NIH","CA","IA","NG","OX"],
  "Process Gas":  ["H2","FG","LG","VG","RG","RD","OG","EOG"],
  Steam:          ["ST","CS"],
  Chemical:       ["BIO","CI","COA","DIS","FL","SI"],
};
const FLUID_COLORS = {
  FW:"#2563eb", PW:"#3b82f6", WM:"#60a5fa", GW:"#22d3ee", SW:"#0ea5e9",
  WW:"#64748b", MCWS:"#06b6d4", MCR:"#0891b2", PWS:"#818cf8", PWR:"#6366f1",
  PCC:"#a78bfa", SL:"#92400e", TW:"#b45309", NI:"#7c3aed", NIH:"#6d28d9",
  CA:"#16a34a", IA:"#15803d", NG:"#d97706", OX:"#dc2626", H2:"#9333ea",
  FG:"#c2410c", LG:"#b45309", VG:"#6b7280", RG:"#475569", RD:"#ef4444",
  OG:"#ea580c", EOG:"#f97316", ST:"#94a3b8", CS:"#cbd5e1",
  BIO:"#65a30d", CI:"#0891b2", COA:"#78350f", DIS:"#059669", FL:"#dc2626", SI:"#7c3aed",
};
const getFluidColor = (sub) => FLUID_COLORS[sub] || "#94a3b8";

const EQUIP_DEFAULTS = {
  Tank:             { capacity:"100 m³",      material:"SS304",     designP:"5 Bar g",   designT:"80 ℃"  },
  Pump:             { capacity:"50 m³/h",     material:"CI",        designP:"10 Bar g",  designT:"60 ℃"  },
  Pond:             { capacity:"500 m³",      material:"Concrete",  designP:"0.5 Bar g", designT:"40 ℃"  },
  "Heat Exchanger": { capacity:"500 kW",      material:"SS316",     designP:"16 Bar g",  designT:"200 ℃" },
  Filter:           { capacity:"20 m³/h",     material:"SS304",     designP:"6 Bar g",   designT:"80 ℃"  },
  Hopper:           { capacity:"10 m³",       material:"MS",        designP:"1 Bar g",   designT:"50 ℃"  },
  Decanter:         { capacity:"15 m³/h",     material:"SS316",     designP:"4 Bar g",   designT:"60 ℃"  },
  "Cooling Tower":  { capacity:"1000 kW",     material:"FRP",       designP:"3 Bar g",   designT:"45 ℃"  },
  Clarifier:        { capacity:"200 m³/h",    material:"MS+Coated", designP:"1 Bar g",   designT:"40 ℃"  },
  Classifier:       { capacity:"50 t/h",      material:"MS",        designP:"1 Bar g",   designT:"60 ℃"  },
  "Feed Box":       { capacity:"5 m³",        material:"MS",        designP:"2 Bar g",   designT:"50 ℃"  },
  "Chemical Dosing":{ capacity:"100 L/h",     material:"PP",        designP:"5 Bar g",   designT:"40 ℃"  },
  Scrubber:         { capacity:"10000 Nm³/h", material:"FRP",       designP:"0.5 Bar g", designT:"60 ℃"  },
  "Bag Filter":     { capacity:"5000 Nm³/h",  material:"CS",        designP:"0.3 Bar g", designT:"180 ℃" },
  Reactor:          { capacity:"50 m³",       material:"SS316L",    designP:"10 Bar g",  designT:"150 ℃" },
  "Feed Bin":       { capacity:"20 m³",       material:"MS",        designP:"1 Bar g",   designT:"50 ℃"  },
  "Gas Duct":       { capacity:"20000 Nm³/h", material:"MS",        designP:"0.5 Bar g", designT:"300 ℃" },
  "Steel Structure":{ capacity:"-",           material:"A36 Steel", designP:"-",         designT:"-"      },
};

// ─────────────────────────────────────────────────────────────
// ID GENERATOR
// ─────────────────────────────────────────────────────────────
let _idCnt = 1;
const uid = (prefix="n") => `${prefix}_${Date.now()}_${_idCnt++}`;

// ─────────────────────────────────────────────────────────────
// MINI SVG ICONS
// ─────────────────────────────────────────────────────────────
const ICONS = {
  Tank:           <><rect x="3" y="5" width="18" height="14" rx="2" fill="none" stroke="currentColor" strokeWidth="1.5"/><line x1="3" y1="17" x2="21" y2="17" stroke="currentColor" strokeWidth="1.5"/></>,
  Pump:           <><circle cx="12" cy="12" r="8" fill="none" stroke="currentColor" strokeWidth="1.5"/><polygon points="12,5 19,12 12,19 5,12" fill="none" stroke="currentColor" strokeWidth="1"/></>,
  "Heat Exchanger":<><rect x="3" y="7" width="18" height="10" rx="1" fill="none" stroke="currentColor" strokeWidth="1.5"/>{[7,11,15].map(x=><line key={x} x1={x} y1="7" x2={x} y2="17" stroke="currentColor" strokeWidth="1"/>)}</>,
  Filter:         <><polygon points="3,3 21,3 15,21 9,21" fill="none" stroke="currentColor" strokeWidth="1.5"/></>,
  "Cooling Tower":<><path d="M4,3 L20,3 L18,21 L6,21 Z" fill="none" stroke="currentColor" strokeWidth="1.5"/><line x1="12" y1="3" x2="12" y2="21" stroke="currentColor" strokeWidth="1"/></>,
  Brench:         <><circle cx="12" cy="12" r="4" fill="none" stroke="currentColor" strokeWidth="1.5"/><line x1="2" y1="12" x2="8" y2="12" stroke="currentColor" strokeWidth="1.5"/><line x1="16" y1="12" x2="22" y2="12" stroke="currentColor" strokeWidth="1.5"/><line x1="12" y1="2" x2="12" y2="8" stroke="currentColor" strokeWidth="1.5"/></>,
};
const EquipSVG = ({ type, size=22 }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" style={{ display:"block", flexShrink:0 }}>
    {ICONS[type] || (
      <><rect x="3" y="3" width="18" height="18" rx="3" fill="none" stroke="currentColor" strokeWidth="1.5"/>
        <text x="12" y="15" textAnchor="middle" fontSize="8" fill="currentColor" fontWeight="600">{type.substring(0,3).toUpperCase()}</text></>
    )}
  </svg>
);

const DIRS = [Position.Top, Position.Bottom, Position.Left, Position.Right];
const DIR_ID = {
  [Position.Top]:"top", [Position.Bottom]:"bottom",
  [Position.Left]:"left", [Position.Right]:"right",
};

// ─────────────────────────────────────────────────────────────
// AREA INLET/OUTLET COMPUTATION UTILITY
// Returns { inlets: string[], outlets: string[] } for one area node
// ─────────────────────────────────────────────────────────────
const computeAreaIO = (areaNode, allNodes, allEdges) => {
  const ax = areaNode.position.x;
  const ay = areaNode.position.y;
  const aw = areaNode.style?.width  || areaNode.width  || 240;
  const ah = areaNode.style?.height || areaNode.height || 160;

  // Which non-area nodes are inside this area's bounding box?
  const insideIds = new Set(
    allNodes
      .filter(n => {
        if (n.type === "area") return false;
        const nx = n.position.x;
        const ny = n.position.y;
        return nx >= ax && ny >= ay && nx <= ax + aw && ny <= ay + ah;
      })
      .map(n => n.id)
  );

  if (insideIds.size === 0) return { inlets:[], outlets:[] };

  // Edges where exactly one endpoint is inside → boundary-crossing
  const inlets  = [];
  const outlets = [];

  allEdges.forEach(e => {
    const srcIn = insideIds.has(e.source);
    const tgtIn = insideIds.has(e.target);
    if (srcIn === tgtIn) return; // both inside or both outside → skip

    const d   = e.data || {};
    const sub  = d.fluidSub  || d.lineType || "—";
    const size = d.size ? ` ${d.size}` : "";
    const label = `${sub}${size}`;

    if (tgtIn) {
      // flow goes INTO the area
      inlets.push(label);
    } else {
      // flow goes OUT of the area
      outlets.push(label);
    }
  });

  return { inlets, outlets };
};

// ─────────────────────────────────────────────────────────────
// AREA NODE  — shows title + auto IO summary + manual notes
// ─────────────────────────────────────────────────────────────
const AreaNode = memo(({ data, selected }) => {
  const c = AREA_COLORS[data.areaType] || AREA_COLORS.Plant;
  const inlets  = data.autoInlets  || [];
  const outlets = data.autoOutlets || [];
  const hasIO   = inlets.length > 0 || outlets.length > 0;

  return (
    <div style={{
      background:c.bg,
      border:`2px dashed ${selected?"#f59e0b":c.border}`,
      borderRadius:10, width:"100%", height:"100%",
      position:"relative", boxSizing:"border-box", overflow:"hidden",
    }}>
      <NodeResizer minWidth={180} minHeight={120} isVisible={selected}/>

      {/* Title bar */}
      <div style={{
        position:"absolute", top:0, left:0, right:0,
        padding:"5px 10px 4px",
        background: selected ? "rgba(245,158,11,0.12)" : "rgba(255,255,255,0.55)",
        borderBottom:`1px dashed ${c.border}`,
        userSelect:"none", pointerEvents:"none",
      }}>
        <span style={{ fontSize:11, fontWeight:800, color:c.label }}>
          [{data.areaType}]{data.label ? ` ${data.label}` : ""}
        </span>
      </div>

      {/* IO summary — auto computed */}
      {hasIO && (
        <div style={{
          position:"absolute", top:28, left:8, right:8,
          display:"flex", gap:6,
          userSelect:"none", pointerEvents:"none",
        }}>
          {inlets.length > 0 && (
            <div style={{
              flex:1, background:"rgba(255,255,255,0.72)",
              border:`1px solid ${c.border}`, borderRadius:5,
              padding:"4px 6px", fontSize:9, lineHeight:1.6,
            }}>
              <div style={{ fontWeight:800, color:"#1d4ed8", marginBottom:1 }}>▶ INLET</div>
              {inlets.map((s,i) => (
                <div key={i} style={{ color:"#1e293b", display:"flex", alignItems:"center", gap:3 }}>
                  <div style={{ width:6, height:6, borderRadius:"50%", background: getFluidColor(s.split(" ")[0]), flexShrink:0 }}/>
                  {s}
                </div>
              ))}
            </div>
          )}
          {outlets.length > 0 && (
            <div style={{
              flex:1, background:"rgba(255,255,255,0.72)",
              border:`1px solid ${c.border}`, borderRadius:5,
              padding:"4px 6px", fontSize:9, lineHeight:1.6,
            }}>
              <div style={{ fontWeight:800, color:"#dc2626", marginBottom:1 }}>◀ OUTLET</div>
              {outlets.map((s,i) => (
                <div key={i} style={{ color:"#1e293b", display:"flex", alignItems:"center", gap:3 }}>
                  <div style={{ width:6, height:6, borderRadius:"50%", background: getFluidColor(s.split(" ")[0]), flexShrink:0 }}/>
                  {s}
                </div>
              ))}
            </div>
          )}
        </div>
      )}

      {/* Manual notes — always at bottom */}
      {data.summary && (
        <div style={{
          position:"absolute", bottom:6, left:8, right:8,
          fontSize:9, color:c.label, lineHeight:1.4,
          userSelect:"none", pointerEvents:"none",
          whiteSpace:"pre-wrap", opacity:0.85,
        }}>
          {data.summary}
        </div>
      )}
    </div>
  );
});

// ─────────────────────────────────────────────────────────────
// EQUIPMENT NODE
// ─────────────────────────────────────────────────────────────
const EquipmentNode = memo(({ id, data, selected }) => {
  const dirs   = data.handles || ["top","bottom","left","right"];
  const posMap = { top:Position.Top, bottom:Position.Bottom, left:Position.Left, right:Position.Right };

  const sideCount = {};
  dirs.forEach(d => { sideCount[d] = (sideCount[d]||0)+1; });
  const sideCursor = {};
  const handleList = dirs.map((dir, i) => {
    sideCursor[dir] = (sideCursor[dir]||0);
    const idx   = sideCursor[dir]++;
    const total = sideCount[dir];
    const pct   = total===1 ? 50 : 20 + (idx/(total-1))*60;
    return { dir, pid:`${dir}_${i}`, pct };
  });

  const getStyle = (dir, pct) => {
    const base = { width:9, height:9, borderRadius:"50%", background:"#3b82f6", border:"2px solid #fff", zIndex:10 };
    if (dir==="top"||dir==="bottom") return { ...base, left:`${pct}%`, transform:"translateX(-50%)" };
    return { ...base, top:`${pct}%`, transform:"translateY(-50%)" };
  };

  return (
    <div style={{
      background:selected?"#eff6ff":"#ffffff",
      border:`${selected?2:1.5}px solid ${selected?"#3b82f6":"#cbd5e1"}`,
      borderRadius:8, minWidth:110, minHeight:64,
      display:"flex", flexDirection:"column", alignItems:"center", justifyContent:"center",
      padding:"8px 12px", fontSize:11, cursor:"default",
      position:"relative", userSelect:"none", boxSizing:"border-box",
    }}>
      <NodeResizer minWidth={80} minHeight={50} isVisible={selected} handleStyle={{ width:8, height:8 }}/>
      {handleList.map(({ dir, pid, pct }) => (
        <Handle key={pid} type="source" position={posMap[dir]} id={pid} style={getStyle(dir,pct)}/>
      ))}
      <div style={{ color:"#3b82f6", marginBottom:4 }}><EquipSVG type={data.equipType} size={26}/></div>
      <div style={{ fontWeight:700, color:"#0f172a", fontSize:11, textAlign:"center", lineHeight:1.3 }}>
        {data.itemNo||data.equipType}
      </div>
      {data.label && data.label!==(data.itemNo||data.equipType) && (
        <div style={{ color:"#64748b", fontSize:10, textAlign:"center" }}>{data.label}</div>
      )}
    </div>
  );
});

// ─────────────────────────────────────────────────────────────
// BRENCH NODE
// ─────────────────────────────────────────────────────────────
const BrenchNode = memo(({ data, selected }) => (
  <div style={{
    width:18, height:18, borderRadius:"50%",
    background:selected?"#f59e0b":"#334155",
    border:`2px solid ${selected?"#b45309":"#1e293b"}`,
    position:"relative",
  }}>
    {DIRS.map(pos=>(
      <Handle key={pos} type="source" position={pos} id={DIR_ID[pos]}
        style={{ width:7, height:7, borderRadius:"50%", background:"#94a3b8", border:"1px solid #fff" }}/>
    ))}
  </div>
));

// ─────────────────────────────────────────────────────────────
// INSTRUMENT NODE
// ─────────────────────────────────────────────────────────────
const InstrumentNode = memo(({ data, selected }) => (
  <div style={{
    width:54, height:54, borderRadius:"50%",
    background:selected?"#faf5ff":"#fafafa",
    border:`${selected?2:1.5}px solid ${selected?"#a855f7":"#c4b5fd"}`,
    display:"flex", flexDirection:"column", alignItems:"center", justifyContent:"center",
    fontSize:10, position:"relative", userSelect:"none",
  }}>
    {DIRS.map(pos=>(
      <Handle key={pos} type="source" position={pos} id={DIR_ID[pos]}
        style={{ width:7, height:7, background:"#a855f7", border:"1px solid #fff" }}/>
    ))}
    <div style={{ fontWeight:800, color:"#7c3aed", fontSize:12, lineHeight:1 }}>
      {(data.instrCategory||"?").charAt(0)}{(data.instrType||"?").charAt(0)}
    </div>
    <div style={{ color:"#94a3b8", fontSize:9 }}>{data.itemNo||""}</div>
  </div>
));

// ─────────────────────────────────────────────────────────────
// PIPE EDGE
// ─────────────────────────────────────────────────────────────
const PipeEdge = ({
  id, sourceX, sourceY, targetX, targetY,
  sourcePosition, targetPosition, data, selected,
}) => {
  const [edgePath] = getSmoothStepPath({
    sourceX, sourceY, sourcePosition,
    targetX, targetY, targetPosition,
    borderRadius:4,
  });
  const color  = data?.fluidSub ? getFluidColor(data.fluidSub) : data?.lineType==="Duct"?"#475569":"#94a3b8";
  const sw     = data?.lineType==="Duct" ? 3.5 : 2;
  const sdash  = data?.lineType==="Conveyor" ? "7,3" : "none";
  const stroke = selected?"#f59e0b":color;
  const mkId   = `mk_${id}`;
  return (
    <g>
      <defs>
        <marker id={mkId} markerWidth="5" markerHeight="4" refX="4.5" refY="2" orient="auto" markerUnits="strokeWidth">
          <polygon points="0 0, 5 2, 0 4" fill={stroke}/>
        </marker>
      </defs>
      <path d={edgePath} fill="none" stroke="transparent" strokeWidth={14} style={{ cursor:"pointer" }}/>
      <path d={edgePath} fill="none" stroke={stroke} strokeWidth={selected?sw+0.5:sw}
        strokeDasharray={sdash} markerEnd={`url(#${mkId})`} style={{ pointerEvents:"none" }}/>
      {(data?.serialNo||data?.size) && (
        <EdgeLabelRenderer>
          <div style={{
            position:"absolute",
            transform:`translate(-50%,-50%) translate(${(sourceX+targetX)/2}px,${(sourceY+targetY)/2}px)`,
            fontSize:9, background:"rgba(255,255,255,0.88)", padding:"1px 4px",
            borderRadius:3, border:`1px solid ${color}`, color,
            pointerEvents:"none", whiteSpace:"nowrap",
          }}>
            {[data.size,data.serialNo].filter(Boolean).join(" · ")}
          </div>
        </EdgeLabelRenderer>
      )}
    </g>
  );
};

// ─────────────────────────────────────────────────────────────
// TYPE MAPS
// ─────────────────────────────────────────────────────────────
const nodeTypes = { area:AreaNode, equipment:EquipmentNode, instrument:InstrumentNode, brench:BrenchNode };
const edgeTypes = { pipe:PipeEdge };

// ─────────────────────────────────────────────────────────────
// CATALOG SIDEBAR
// ─────────────────────────────────────────────────────────────
const Sidebar = memo(({ onDragStart }) => {
  const [open, setOpen] = useState({ Area:true, Equipment:true, Connection:false, Instrument:false });
  const tog = k => setOpen(p=>({...p,[k]:!p[k]}));
  const iS  = (color="#475569") => ({
    padding:"5px 14px 5px 20px", cursor:"grab",
    borderBottom:"1px solid #f1f5f9", color, fontSize:11,
    userSelect:"none", display:"flex", alignItems:"center", gap:6,
  });
  const hov = {
    onMouseEnter: e=>{ e.currentTarget.style.background="#e0f2fe"; },
    onMouseLeave: e=>{ e.currentTarget.style.background=""; },
  };
  const Sec = ({ title, cat, children }) => (
    <div>
      <div onClick={()=>tog(cat)} style={{
        padding:"7px 14px", fontWeight:600, color:"#334155", cursor:"pointer",
        background:"#f1f5f9", borderBottom:"1px solid #e2e8f0",
        display:"flex", justifyContent:"space-between", userSelect:"none", fontSize:12,
      }}>
        <span>{title}</span><span style={{ fontSize:10 }}>{open[cat]?"▲":"▼"}</span>
      </div>
      {open[cat] && children}
    </div>
  );
  return (
    <div style={{ width:195, background:"#f8fafc", borderRight:"1px solid #e2e8f0", overflowY:"auto", flexShrink:0 }}>
      <div style={{ padding:"10px 14px 7px", fontWeight:800, fontSize:13, color:"#0f172a", borderBottom:"1px solid #e2e8f0", letterSpacing:0.5 }}>Catalog</div>
      <Sec title="Area" cat="Area">
        {AREA_TYPES.map(at=>{ const c=AREA_COLORS[at]; return (
          <div key={at} draggable onDragStart={e=>onDragStart(e,"area",at)}
            style={{ ...iS(c.label), background:c.bg }} {...hov}>▭ {at}</div>
        );})}
      </Sec>
      <Sec title="Equipment" cat="Equipment">
        {EQUIPMENT_LIST.map(eq=>(
          <div key={eq} draggable onDragStart={e=>onDragStart(e,"equipment",eq)} style={iS()} {...hov}>
            <EquipSVG type={eq} size={14}/>{eq}
          </div>
        ))}
      </Sec>
      <Sec title="Connection" cat="Connection">
        {CONNECTION_LIST.map(c=>(
          <div key={c} draggable onDragStart={e=>onDragStart(e,"connection",c)} style={iS("#1d4ed8")} {...hov}>
            {c==="Piping"?"━━ ":c==="Duct"?"▬▬ ":"⊕ "}{c}
          </div>
        ))}
        {CONVEYOR_LIST.map(c=>(
          <div key={c} draggable onDragStart={e=>onDragStart(e,"conveyor",c)} style={iS("#78350f")} {...hov}>╌╌ {c}</div>
        ))}
      </Sec>
      <Sec title="Instrument" cat="Instrument">
        {Object.keys(INSTRUMENT_CATS).map(cat=>
          INSTR_TYPES.map(t=>(
            <div key={`${cat}-${t}`} draggable onDragStart={e=>onDragStart(e,"instrument",`${cat}|${t}`)} style={iS("#7c3aed")} {...hov}>
              ⊙ {cat} {t}
            </div>
          ))
        )}
      </Sec>
    </div>
  );
});

// ─────────────────────────────────────────────────────────────
// INSPECTOR
// ─────────────────────────────────────────────────────────────
const Inspector = memo(({ sel, nodes, edges, onUpdateNode, onUpdateEdge, onDeleteSel, onAddHandle }) => {
  const [reqText, setReqText] = useState("");
  const [reqWho,  setReqWho]  = useState("");
  const [reqDate, setReqDate] = useState(new Date().toISOString().slice(0,10));
  const [tab, setTab]         = useState("spec");

  const L = { fontSize:11, color:"#64748b", marginBottom:2, display:"block", fontWeight:500 };
  const I = { width:"100%", padding:"4px 7px", border:"1px solid #e2e8f0", borderRadius:4, fontSize:12, boxSizing:"border-box", marginBottom:7 };
  const S = { ...I, background:"#fff" };

  if (!sel) return (
    <div style={{ width:270, background:"#fff", borderLeft:"1px solid #e2e8f0", padding:16, color:"#94a3b8", fontSize:12, flexShrink:0 }}>
      <div style={{ fontWeight:800, fontSize:13, color:"#0f172a", marginBottom:10 }}>Inspector</div>
      노드 또는 엣지를 선택하세요.
    </div>
  );

  const isNode = !!sel.position;
  const d = sel.data||{};
  const upN = (k,v) => onUpdateNode(sel.id,{...d,[k]:v});
  const upE = (k,v) => onUpdateEdge(sel.id,{...d,[k]:v});

  const addReq = () => {
    if (!reqText.trim()) return;
    const reqs=[...(d.requirements||[]),{id:Date.now(),text:reqText,who:reqWho,date:reqDate}];
    onUpdateNode(sel.id,{...d,requirements:reqs});
    setReqText(""); setReqWho("");
  };

  const connEdges = edges.filter(e=>e.source===sel.id||e.target===sel.id);
  const ifaceList = connEdges.map(e=>{
    const other = e.source===sel.id ? nodes.find(n=>n.id===e.target) : nodes.find(n=>n.id===e.source);
    return { edgeId:e.id, dir:e.source===sel.id?"→":"←",
      otherLabel:other?.data?.itemNo||other?.data?.label||other?.id||"?", ...e.data };
  });

  const TB = ({ name, label }) => (
    <button onClick={()=>setTab(name)} style={{
      flex:1, padding:"5px 0", border:"none", cursor:"pointer", fontSize:11,
      background:tab===name?"#eff6ff":"#f8fafc",
      color:tab===name?"#1d4ed8":"#64748b",
      fontWeight:tab===name?700:400,
      borderBottom:tab===name?"2px solid #3b82f6":"2px solid transparent",
    }}>{label}</button>
  );

  return (
    <div style={{ width:270, background:"#fff", borderLeft:"1px solid #e2e8f0", overflowY:"auto", flexShrink:0, display:"flex", flexDirection:"column" }}>
      <div style={{ padding:"10px 12px 7px", borderBottom:"1px solid #e2e8f0", display:"flex", justifyContent:"space-between", alignItems:"center", flexShrink:0 }}>
        <span style={{ fontWeight:800, fontSize:13, color:"#0f172a" }}>Inspector</span>
        <button onClick={onDeleteSel} style={{ background:"#fee2e2", color:"#dc2626", border:"none", borderRadius:4, padding:"2px 9px", cursor:"pointer", fontSize:11 }}>Delete</button>
      </div>

      {isNode && (
        <div style={{ display:"flex", borderBottom:"1px solid #e2e8f0", flexShrink:0 }}>
          <TB name="spec" label="Spec"/>
          <TB name="req"  label="Req."/>
          <TB name="iface" label="I/F"/>
          {sel.type==="area" && <TB name="io" label="IO"/>}
        </div>
      )}

      <div style={{ padding:"10px 12px", overflowY:"auto", flex:1 }}>

        {/* ── SPEC ── */}
        {(tab==="spec"||!isNode) && (
          <>
            {isNode && sel.type==="area" && (
              <>
                <label style={L}>Area Type</label>
                <select style={S} value={d.areaType||"Plant"} onChange={e=>upN("areaType",e.target.value)}>
                  {AREA_TYPES.map(a=><option key={a}>{a}</option>)}
                </select>
                <label style={L}>Label</label>
                <input style={I} value={d.label||""} onChange={e=>upN("label",e.target.value)}/>
                <label style={L}>Notes (직접입력)</label>
                <textarea style={{ ...I, height:70, resize:"vertical" }} value={d.summary||""} onChange={e=>upN("summary",e.target.value)}/>
              </>
            )}
            {isNode && sel.type==="equipment" && (
              <>
                <label style={L}>Item No</label>
                <input style={I} value={d.itemNo||""} onChange={e=>upN("itemNo",e.target.value)} placeholder="e.g. T-101"/>
                <label style={L}>Description</label>
                <input style={I} value={d.label||""} onChange={e=>upN("label",e.target.value)}/>
                <div style={{ fontWeight:600, fontSize:11, color:"#334155", margin:"6px 0 5px", borderTop:"1px solid #f1f5f9", paddingTop:5 }}>Engineering Spec</div>
                {[["material","Material"],["capacity","Capacity"],["designP","Design Pressure"],["designT","Design Temp"]].map(([k,l])=>(
                  <React.Fragment key={k}>
                    <label style={L}>{l}</label>
                    <input style={I} value={d[k]||""} onChange={e=>upN(k,e.target.value)}/>
                  </React.Fragment>
                ))}
                <div style={{ fontWeight:600, fontSize:11, color:"#334155", margin:"6px 0 5px", borderTop:"1px solid #f1f5f9", paddingTop:5 }}>Port Management</div>
                <div style={{ display:"flex", flexWrap:"wrap", gap:5, marginBottom:6 }}>
                  {["top","bottom","left","right"].map(dir=>(
                    <button key={dir} onClick={()=>onAddHandle(sel.id,dir)}
                      style={{ background:"#eff6ff", border:"1px solid #bfdbfe", color:"#1d4ed8", borderRadius:4, padding:"2px 8px", cursor:"pointer", fontSize:11 }}>
                      + {dir}
                    </button>
                  ))}
                </div>
                <div style={{ fontSize:10, color:"#94a3b8" }}>Ports: {(d.handles||[]).join(", ")}</div>
              </>
            )}
            {isNode && sel.type==="instrument" && (
              <>
                <label style={L}>Item No</label>
                <input style={I} value={d.itemNo||""} onChange={e=>upN("itemNo",e.target.value)}/>
                <label style={L}>Category</label>
                <select style={S} value={d.instrCategory||""} onChange={e=>upN("instrCategory",e.target.value)}>
                  <option value="">Select</option>
                  {Object.keys(INSTRUMENT_CATS).map(c=><option key={c}>{c}</option>)}
                </select>
                <label style={L}>Type</label>
                <select style={S} value={d.instrType||""} onChange={e=>upN("instrType",e.target.value)}>
                  <option value="">Select</option>
                  {INSTR_TYPES.map(t=><option key={t}>{t}</option>)}
                </select>
              </>
            )}
            {isNode && sel.type==="brench" && (
              <div style={{ fontSize:12, color:"#64748b" }}>Brench (Junction) node.<br/>배관 분기점으로 사용합니다.</div>
            )}
            {!isNode && (
              <>
                <label style={L}>Line Type</label>
                <select style={S} value={d.lineType||"Piping"} onChange={e=>upE("lineType",e.target.value)}>
                  {[...CONNECTION_LIST,...CONVEYOR_LIST].map(c=><option key={c}>{c}</option>)}
                </select>
                {(d.lineType==="Piping"||!d.lineType) && (
                  <>
                    <label style={L}>Serial No (Line No)</label>
                    <input style={I} value={d.serialNo||""} onChange={e=>upE("serialNo",e.target.value)} placeholder="e.g. L-101"/>
                    <label style={L}>Size (DN)</label>
                    <input style={I} value={d.size||""} onChange={e=>upE("size",e.target.value)} placeholder="e.g. DN200"/>
                    <label style={L}>Schedule (Spec)</label>
                    <input style={I} value={d.spec||""} onChange={e=>upE("spec",e.target.value)} placeholder="e.g. SCH40"/>
                    <div style={{ fontWeight:600, fontSize:11, color:"#334155", margin:"6px 0 5px", borderTop:"1px solid #f1f5f9", paddingTop:5 }}>Fluid</div>
                    <label style={L}>Primary</label>
                    <select style={S} value={d.fluidPrimary||""} onChange={e=>onUpdateEdge(sel.id,{...d,fluidPrimary:e.target.value,fluidSub:""})}>
                      <option value="">Select</option>
                      {FLUID_PRIMARY.map(f=><option key={f}>{f}</option>)}
                    </select>
                    {d.fluidPrimary && FLUID_SECONDARY[d.fluidPrimary] && (
                      <>
                        <label style={L}>Sub-type</label>
                        <select style={S} value={d.fluidSub||""} onChange={e=>upE("fluidSub",e.target.value)}>
                          <option value="">Select</option>
                          {FLUID_SECONDARY[d.fluidPrimary].map(f=><option key={f}>{f}</option>)}
                        </select>
                      </>
                    )}
                    {d.fluidSub && (
                      <div style={{ display:"flex", alignItems:"center", gap:8, marginTop:4, marginBottom:8 }}>
                        <div style={{ width:28, height:8, borderRadius:4, background:getFluidColor(d.fluidSub) }}/>
                        <span style={{ fontSize:11, color:"#64748b" }}>{d.fluidSub}</span>
                      </div>
                    )}
                  </>
                )}
              </>
            )}
          </>
        )}

        {/* ── REQUIREMENTS ── */}
        {tab==="req" && isNode && (
          <>
            {(d.requirements||[]).map(r=>(
              <div key={r.id} style={{ background:"#fffbeb", border:"1px solid #fde68a", borderRadius:6, padding:"6px 8px", marginBottom:6, fontSize:11 }}>
                <div style={{ display:"flex", justifyContent:"space-between", marginBottom:2 }}>
                  <span style={{ fontWeight:700, color:"#92400e" }}>{r.who||"—"}</span>
                  <span style={{ color:"#a16207", fontSize:10 }}>{r.date}</span>
                </div>
                <div style={{ color:"#1c1917", lineHeight:1.4 }}>{r.text}</div>
                <button onClick={()=>{ const rs=(d.requirements||[]).filter(x=>x.id!==r.id); onUpdateNode(sel.id,{...d,requirements:rs}); }}
                  style={{ background:"none", border:"none", color:"#dc2626", cursor:"pointer", fontSize:10, padding:"2px 0", marginTop:2 }}>× remove</button>
              </div>
            ))}
            {!(d.requirements||[]).length && <div style={{ color:"#94a3b8", fontSize:11, marginBottom:10 }}>등록된 요구사항 없음</div>}
            <div style={{ borderTop:"1px solid #f1f5f9", paddingTop:8 }}>
              <textarea style={{ ...I, height:56, resize:"vertical" }} placeholder="Requirement..." value={reqText} onChange={e=>setReqText(e.target.value)}/>
              <input style={I} placeholder="Stakeholder" value={reqWho} onChange={e=>setReqWho(e.target.value)}/>
              <input type="date" style={I} value={reqDate} onChange={e=>setReqDate(e.target.value)}/>
              <button onClick={addReq} style={{ width:"100%", padding:"6px", background:"#1d4ed8", color:"#fff", border:"none", borderRadius:5, cursor:"pointer", fontSize:12, fontWeight:700 }}>
                + ADD Requirement
              </button>
            </div>
          </>
        )}

        {/* ── INTERFACE LIST ── */}
        {tab==="iface" && isNode && (
          <>
            <div style={{ fontWeight:600, fontSize:11, color:"#334155", marginBottom:6 }}>Connected Interfaces ({ifaceList.length})</div>
            {ifaceList.length===0 && <div style={{ color:"#94a3b8", fontSize:11 }}>연결된 인터페이스 없음</div>}
            {ifaceList.map(iface=>(
              <div key={iface.edgeId} style={{ border:"1px solid #e2e8f0", borderRadius:6, padding:"6px 8px", marginBottom:5, fontSize:11 }}>
                <div style={{ display:"flex", justifyContent:"space-between" }}>
                  <span style={{ fontWeight:700, color:"#1d4ed8" }}>{iface.dir} {iface.otherLabel}</span>
                  <span style={{ color:"#64748b" }}>{iface.lineType||"Piping"}</span>
                </div>
                {iface.fluidSub && (
                  <div style={{ display:"flex", alignItems:"center", gap:5, marginTop:3 }}>
                    <div style={{ width:16, height:5, borderRadius:3, background:getFluidColor(iface.fluidSub) }}/>
                    <span style={{ color:"#64748b" }}>{iface.fluidSub} {iface.size||""} {iface.serialNo||""}</span>
                  </div>
                )}
              </div>
            ))}
          </>
        )}

        {/* ── IO SUMMARY TAB (Area only) ── */}
        {tab==="io" && isNode && sel.type==="area" && (
          <>
            <div style={{ fontWeight:700, fontSize:11, color:"#1d4ed8", marginBottom:8 }}>
              ▶ INLET ({(d.autoInlets||[]).length})
            </div>
            {(d.autoInlets||[]).length===0
              ? <div style={{ color:"#94a3b8", fontSize:11, marginBottom:10 }}>없음 (Area 안의 노드로 들어오는 연결 없음)</div>
              : (d.autoInlets||[]).map((s,i)=>(
                <div key={i} style={{ display:"flex", alignItems:"center", gap:8, padding:"4px 8px", background:"#eff6ff", borderRadius:5, marginBottom:4, fontSize:11 }}>
                  <div style={{ width:10, height:10, borderRadius:"50%", background:getFluidColor(s.split(" ")[0]), flexShrink:0 }}/>
                  <span style={{ fontWeight:600, color:"#1e293b" }}>{s}</span>
                </div>
              ))
            }
            <div style={{ fontWeight:700, fontSize:11, color:"#dc2626", marginBottom:8, marginTop:10 }}>
              ◀ OUTLET ({(d.autoOutlets||[]).length})
            </div>
            {(d.autoOutlets||[]).length===0
              ? <div style={{ color:"#94a3b8", fontSize:11, marginBottom:10 }}>없음 (Area 밖으로 나가는 연결 없음)</div>
              : (d.autoOutlets||[]).map((s,i)=>(
                <div key={i} style={{ display:"flex", alignItems:"center", gap:8, padding:"4px 8px", background:"#fff1f2", borderRadius:5, marginBottom:4, fontSize:11 }}>
                  <div style={{ width:10, height:10, borderRadius:"50%", background:getFluidColor(s.split(" ")[0]), flexShrink:0 }}/>
                  <span style={{ fontWeight:600, color:"#1e293b" }}>{s}</span>
                </div>
              ))
            }
            <div style={{ marginTop:12, padding:"8px", background:"#f8fafc", borderRadius:6, fontSize:10, color:"#64748b", lineHeight:1.6 }}>
              💡 Area 경계 안에 Equipment/Instrument를 배치하고 Pipe를 연결하면 자동으로 집계됩니다.
            </div>
          </>
        )}

      </div>
    </div>
  );
});

// ─────────────────────────────────────────────────────────────
// CONNECTION TYPE MODAL
// ─────────────────────────────────────────────────────────────
const ConnModal = ({ onConfirm, onCancel }) => {
  const [lt, setLt] = useState("Piping");
  return (
    <div style={{ position:"fixed", inset:0, background:"rgba(0,0,0,0.45)", display:"flex", alignItems:"center", justifyContent:"center", zIndex:9999 }}>
      <div style={{ background:"#fff", borderRadius:10, padding:24, minWidth:260, boxShadow:"0 8px 30px rgba(0,0,0,0.18)" }}>
        <div style={{ fontWeight:800, fontSize:15, marginBottom:14, color:"#0f172a" }}>New Connection</div>
        <label style={{ fontSize:12, color:"#64748b", display:"block", marginBottom:4 }}>Line Type</label>
        <select value={lt} onChange={e=>setLt(e.target.value)}
          style={{ width:"100%", padding:"6px 8px", border:"1px solid #e2e8f0", borderRadius:6, fontSize:13, marginBottom:16 }}>
          {[...CONNECTION_LIST,...CONVEYOR_LIST].map(c=><option key={c}>{c}</option>)}
        </select>
        <div style={{ display:"flex", gap:8 }}>
          <button onClick={()=>onConfirm(lt)} style={{ flex:1, background:"#1d4ed8", color:"#fff", border:"none", borderRadius:6, padding:"8px", cursor:"pointer", fontWeight:700 }}>Create</button>
          <button onClick={onCancel} style={{ flex:1, background:"#f1f5f9", color:"#334155", border:"1px solid #e2e8f0", borderRadius:6, padding:"8px", cursor:"pointer" }}>Cancel</button>
        </div>
      </div>
    </div>
  );
};

// ─────────────────────────────────────────────────────────────
// MAIN CANVAS
// ─────────────────────────────────────────────────────────────
const CanvasInner = () => {
  const wrapRef = useRef(null);
  const { screenToFlowPosition } = useReactFlow();
  const [nodes, setNodes, onNodesChange] = useNodesState([]);
  const [edges, setEdges, onEdgesChange] = useEdgesState([]);
  const [sel,   setSel]   = useState(null);
  const [modal, setModal] = useState(false);
  const connRef = useRef(null);
  const fileRef = useRef(null);

  // ── AUTO-COMPUTE AREA IO whenever nodes or edges change ────
  useEffect(() => {
    const areaNodes = nodes.filter(n => n.type === "area");
    if (areaNodes.length === 0) return;

    let changed = false;
    const updated = nodes.map(n => {
      if (n.type !== "area") return n;
      const { inlets, outlets } = computeAreaIO(n, nodes, edges);
      const same =
        JSON.stringify(n.data.autoInlets)  === JSON.stringify(inlets) &&
        JSON.stringify(n.data.autoOutlets) === JSON.stringify(outlets);
      if (same) return n;
      changed = true;
      return { ...n, data: { ...n.data, autoInlets: inlets, autoOutlets: outlets } };
    });

    if (changed) {
      setNodes(updated);
    }
  }, [edges, nodes.map(n=>n.position.x+","+n.position.y).join("|")]); // re-run on position or edge changes

  // DRAG FROM SIDEBAR
  const onDragStart = useCallback((e,cat,sub)=>{
    e.dataTransfer.setData("mbse/cat",cat);
    e.dataTransfer.setData("mbse/sub",sub);
    e.dataTransfer.effectAllowed="move";
  },[]);
  const onDragOver = useCallback(e=>{ e.preventDefault(); e.dataTransfer.dropEffect="move"; },[]);

  const onDrop = useCallback(e=>{
    e.preventDefault();
    const cat = e.dataTransfer.getData("mbse/cat");
    const sub = e.dataTransfer.getData("mbse/sub");
    if(!cat) return;
    const pos = screenToFlowPosition({ x:e.clientX, y:e.clientY });
    if(cat==="area"){
      setNodes(ns=>[...ns,{ id:uid("area"), type:"area", position:pos,
        style:{ width:260, height:180 },
        data:{ areaType:sub, label:"", summary:"", requirements:[], autoInlets:[], autoOutlets:[] }, zIndex:-1 }]);
    } else if(cat==="equipment"){
      const def=EQUIP_DEFAULTS[sub]||{};
      setNodes(ns=>[...ns,{ id:uid("eq"), type:"equipment", position:pos,
        data:{ equipType:sub, itemNo:"", label:sub,
               handles:["top","bottom","left","right"], requirements:[], ...def } }]);
    } else if(cat==="instrument"){
      const [ic,it]=sub.split("|");
      setNodes(ns=>[...ns,{ id:uid("ins"), type:"instrument", position:pos,
        data:{ instrCategory:ic, instrType:it, itemNo:"", requirements:[] } }]);
    } else if(cat==="connection"){
      setNodes(ns=>[...ns,{ id:uid("br"), type:"brench", position:pos, data:{ _hint:sub } }]);
    }
  },[screenToFlowPosition, setNodes]);

  const onConnect = useCallback(params=>{
    connRef.current=params; setModal(true);
  },[]);

  const confirmConn = useCallback(lineType=>{
    const p=connRef.current; if(!p) return;
    setEdges(es=>addEdge({ ...p, id:uid("e"), type:"pipe", data:{ lineType } }, es));
    setModal(false); connRef.current=null;
  },[setEdges]);

  const onNodeClick  = useCallback((_,n)=>setSel(n),[]);
  const onEdgeClick  = useCallback((_,e)=>setSel(e),[]);
  const onPaneClick  = useCallback(()=>setSel(null),[]);

  const onUpdateNode = useCallback((id,newData)=>{
    setNodes(ns=>ns.map(n=>n.id===id?{...n,data:newData}:n));
    setSel(prev=>prev&&prev.id===id?{...prev,data:newData}:prev);
  },[setNodes]);

  const onUpdateEdge = useCallback((id,newData)=>{
    setEdges(es=>es.map(e=>e.id===id?{...e,data:newData}:e));
    setSel(prev=>prev&&prev.id===id?{...prev,data:newData}:prev);
  },[setEdges]);

  const onAddHandle = useCallback((nodeId,dir)=>{
    setNodes(ns=>ns.map(n=>{
      if(n.id!==nodeId) return n;
      return { ...n, data:{ ...n.data, handles:[...(n.data.handles||[]),dir] } };
    }));
  },[setNodes]);

  const onDeleteSel = useCallback(()=>{
    if(!sel) return;
    if(sel.position){ setNodes(ns=>ns.filter(n=>n.id!==sel.id)); setEdges(es=>es.filter(e=>e.source!==sel.id&&e.target!==sel.id)); }
    else { setEdges(es=>es.filter(e=>e.id!==sel.id)); }
    setSel(null);
  },[sel,setNodes,setEdges]);

  useEffect(()=>{
    const fn=e=>{
      if((e.key==="Delete"||e.key==="Backspace")&&sel){
        const tag=document.activeElement?.tagName?.toLowerCase();
        if(tag==="input"||tag==="textarea"||tag==="select") return;
        onDeleteSel();
      }
    };
    window.addEventListener("keydown",fn);
    return()=>window.removeEventListener("keydown",fn);
  },[sel,onDeleteSel]);

  const onExport=()=>{
    const blob=new Blob([JSON.stringify({nodes,edges},null,2)],{type:"application/json"});
    const a=document.createElement("a");
    a.href=URL.createObjectURL(blob);
    a.download=`MBSE_${new Date().toISOString().slice(0,10)}.json`;
    a.click();
  };
  const onImport=e=>{
    const f=e.target.files[0]; if(!f) return;
    const r=new FileReader();
    r.onload=ev=>{ try{ const p=JSON.parse(ev.target.result); if(p.nodes)setNodes(p.nodes); if(p.edges)setEdges(p.edges); }catch{alert("Invalid JSON");} };
    r.readAsText(f); e.target.value="";
  };

  useEffect(()=>{
    if(!sel?.position) return;
    const fresh=nodes.find(n=>n.id===sel.id);
    if(fresh&&fresh!==sel) setSel(fresh);
  },[nodes]); // eslint-disable-line

  return (
    <>
    <style>{HANDLE_CSS}</style>
    <div style={{ display:"flex", height:"100vh", width:"100vw", fontFamily:"'Segoe UI',sans-serif", background:"#f8fafc", overflow:"hidden" }}>
      <Sidebar onDragStart={onDragStart}/>
      <div style={{ flex:1, display:"flex", flexDirection:"column", minWidth:0 }}>
        <div style={{ height:44, background:"#0f172a", display:"flex", alignItems:"center", padding:"0 16px", gap:10, flexShrink:0 }}>
          <span style={{ color:"#f1f5f9", fontWeight:800, fontSize:14, marginRight:8, letterSpacing:0.5 }}>⬡ MBSE Interface Master</span>
          <button onClick={onExport} style={{ background:"#1e293b", color:"#94a3b8", border:"1px solid #334155", borderRadius:5, padding:"3px 12px", cursor:"pointer", fontSize:12 }}>⬇ Export</button>
          <button onClick={()=>fileRef.current?.click()} style={{ background:"#1e293b", color:"#94a3b8", border:"1px solid #334155", borderRadius:5, padding:"3px 12px", cursor:"pointer", fontSize:12 }}>⬆ Import</button>
          <input ref={fileRef} type="file" accept=".json" style={{ display:"none" }} onChange={onImport}/>
          <span style={{ marginLeft:"auto", color:"#475569", fontSize:11 }}>{nodes.length} nodes · {edges.length} edges</span>
        </div>
        <div style={{ flex:1, display:"flex", minHeight:0 }}>
          <div ref={wrapRef} style={{ flex:1, minWidth:0 }} onDragOver={onDragOver} onDrop={onDrop}>
            <ReactFlow
              nodes={nodes} edges={edges}
              onNodesChange={onNodesChange} onEdgesChange={onEdgesChange}
              onConnect={onConnect}
              onNodeClick={onNodeClick} onEdgeClick={onEdgeClick} onPaneClick={onPaneClick}
              nodeTypes={nodeTypes} edgeTypes={edgeTypes}
              connectionMode={ConnectionMode.Loose}
              connectionLineType="smoothstep"
              connectionLineStyle={{ stroke:"#3b82f6", strokeWidth:2 }}
              defaultEdgeOptions={{ type:"pipe" }}
              fitView snapToGrid snapGrid={[10,10]} deleteKeyCode={null}
            >
              <Controls/>
              <MiniMap nodeColor={n=>n.type==="instrument"?"#a855f7":n.type==="area"?"#93c5fd":"#3b82f6"} maskColor="rgba(0,0,0,0.04)"/>
              <Background variant="dots" gap={20} size={1} color="#cbd5e1"/>
              <Panel position="bottom-left">
                <div style={{ background:"rgba(255,255,255,0.92)", border:"1px solid #e2e8f0", borderRadius:7, padding:"5px 10px", fontSize:10, color:"#64748b" }}>
                  파란 점 드래그로 연결 · Area 안에 노드 배치 후 연결하면 IO 자동 집계 · Del 삭제
                </div>
              </Panel>
            </ReactFlow>
          </div>
          <Inspector sel={sel} nodes={nodes} edges={edges}
            onUpdateNode={onUpdateNode} onUpdateEdge={onUpdateEdge}
            onDeleteSel={onDeleteSel} onAddHandle={onAddHandle}/>
        </div>
      </div>
      {modal && <ConnModal onConfirm={confirmConn} onCancel={()=>{ setModal(false); connRef.current=null; }}/>}
    </div>
    </>
  );
};

export default function App() {
  return <ReactFlowProvider><CanvasInner/></ReactFlowProvider>;
}