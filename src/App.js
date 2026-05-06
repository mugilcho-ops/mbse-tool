// ============================================================
// MBSE Interface Master v8
// 추가: Excel Export / Import (SheetJS)
// React + ReactFlow  |  package.json: "reactflow": "^11.11.4"
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
// SHEETJS DYNAMIC LOADER
// CDN에서 xlsx 라이브러리를 동적으로 불러옴
// ─────────────────────────────────────────────────────────────
const loadXLSX = () => new Promise((resolve, reject) => {
  if (window.XLSX) { resolve(window.XLSX); return; }
  const script = document.createElement("script");
  script.src = "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";
  script.onload  = () => resolve(window.XLSX);
  script.onerror = () => reject(new Error("SheetJS 로드 실패"));
  document.head.appendChild(script);
});

// ─────────────────────────────────────────────────────────────
// EXCEL EXPORT UTILITY
// 시트 3개: Nodes / Edges / Requirements
// ─────────────────────────────────────────────────────────────
const exportToExcel = async (nodes, edges) => {
  const XLSX = await loadXLSX();

  // ── Sheet 1: Nodes ────────────────────────────────────────
  const nodeRows = nodes.map(n => {
    const d = n.data || {};
    return {
      "ID":           n.id,
      "Type":         n.type,
      "Area Type":    d.areaType    || "",
      "Item No":      d.itemNo      || "",
      "Label":        d.label       || "",
      "Equip Type":   d.equipType   || "",
      "Material":     d.material    || "",
      "Capacity":     d.capacity    || "",
      "Design P":     d.designP     || "",
      "Design T":     d.designT     || "",
      "Instr Category": d.instrCategory || "",
      "Instr Type":   d.instrType   || "",
      "Notes":        d.summary     || "",
      "Pos X":        Math.round(n.position?.x || 0),
      "Pos Y":        Math.round(n.position?.y || 0),
    };
  });

  // ── Sheet 2: Edges ────────────────────────────────────────
  const edgeRows = edges.map(e => {
    const d = e.data || {};
    // source/target 노드 이름 찾기
    const srcNode = nodes.find(n => n.id === e.source);
    const tgtNode = nodes.find(n => n.id === e.target);
    const srcLabel = srcNode?.data?.itemNo || srcNode?.data?.label || e.source;
    const tgtLabel = tgtNode?.data?.itemNo || tgtNode?.data?.label || e.target;
    return {
      "Edge ID":      e.id,
      "Line Type":    d.lineType    || "Piping",
      "From (ID)":    e.source,
      "From (Name)":  srcLabel,
      "To (ID)":      e.target,
      "To (Name)":    tgtLabel,
      "Serial No":    d.serialNo   || "",
      "Size (DN)":    d.size       || "",
      "Schedule":     d.spec       || "",
      "Fluid Primary":d.fluidPrimary|| "",
      "Fluid Sub":    d.fluidSub   || "",
      "Line Text":    d.lineText   || "",
    };
  });

  // ── Sheet 3: Requirements ─────────────────────────────────
  const reqRows = [];
  nodes.forEach(n => {
    const d = n.data || {};
    const name = d.itemNo || d.label || n.id;
    (d.requirements || []).forEach(r => {
      reqRows.push({
        "Node ID":      n.id,
        "Node Name":    name,
        "Node Type":    n.type,
        "Req ID":       r.id,
        "Stakeholder":  r.who  || "",
        "Date":         r.date || "",
        "Requirement":  r.text || "",
      });
    });
  });

  // ── 워크북 생성 ──────────────────────────────────────────
  const wb = XLSX.utils.book_new();

  const wsNodes = XLSX.utils.json_to_sheet(nodeRows);
  const wsEdges = XLSX.utils.json_to_sheet(edgeRows);
  const wsReqs  = XLSX.utils.json_to_sheet(reqRows.length ? reqRows : [{ "Note":"등록된 요구사항 없음" }]);

  // 컬럼 너비 자동 조정
  const autoWidth = (ws, rows) => {
    if (!rows.length) return;
    const keys = Object.keys(rows[0]);
    ws["!cols"] = keys.map(k => ({
      wch: Math.max(k.length, ...rows.map(r => String(r[k]||"").length), 10)
    }));
  };
  autoWidth(wsNodes, nodeRows);
  autoWidth(wsEdges, edgeRows);
  autoWidth(wsReqs,  reqRows.length ? reqRows : []);

  XLSX.utils.book_append_sheet(wb, wsNodes, "Nodes");
  XLSX.utils.book_append_sheet(wb, wsEdges, "Edges");
  XLSX.utils.book_append_sheet(wb, wsReqs,  "Requirements");

  const today = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(wb, `MBSE_${today}.xlsx`);
};

// ─────────────────────────────────────────────────────────────
// EXCEL IMPORT UTILITY
// Nodes 시트: Item No / Label / 스펙 업데이트 (위치·연결 유지)
// Edges 시트: Line 속성 업데이트
// Requirements 시트: 요구사항 병합
// ─────────────────────────────────────────────────────────────
const importFromExcel = async (file, nodes, edges, setNodes, setEdges) => {
  const XLSX = await loadXLSX();
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = ev => {
      try {
        const wb = XLSX.read(ev.target.result, { type:"array" });

        // ── Nodes 시트 ──
        const wsN = wb.Sheets["Nodes"];
        if (wsN) {
          const rows = XLSX.utils.sheet_to_json(wsN);
          const updatedNodes = nodes.map(n => {
            const row = rows.find(r => r["ID"] === n.id);
            if (!row) return n;
            return {
              ...n,
              data: {
                ...n.data,
                itemNo:       row["Item No"]      !== undefined ? String(row["Item No"])      : n.data.itemNo,
                label:        row["Label"]         !== undefined ? String(row["Label"])         : n.data.label,
                material:     row["Material"]      !== undefined ? String(row["Material"])      : n.data.material,
                capacity:     row["Capacity"]      !== undefined ? String(row["Capacity"])      : n.data.capacity,
                designP:      row["Design P"]      !== undefined ? String(row["Design P"])      : n.data.designP,
                designT:      row["Design T"]      !== undefined ? String(row["Design T"])      : n.data.designT,
                summary:      row["Notes"]         !== undefined ? String(row["Notes"])         : n.data.summary,
                instrCategory:row["Instr Category"]!== undefined ? String(row["Instr Category"]): n.data.instrCategory,
                instrType:    row["Instr Type"]    !== undefined ? String(row["Instr Type"])    : n.data.instrType,
              }
            };
          });
          setNodes(updatedNodes);
        }

        // ── Edges 시트 ──
        const wsE = wb.Sheets["Edges"];
        if (wsE) {
          const rows = XLSX.utils.sheet_to_json(wsE);
          const updatedEdges = edges.map(e => {
            const row = rows.find(r => r["Edge ID"] === e.id);
            if (!row) return e;
            return {
              ...e,
              data: {
                ...e.data,
                lineType:     row["Line Type"]    !== undefined ? String(row["Line Type"])    : e.data?.lineType,
                serialNo:     row["Serial No"]    !== undefined ? String(row["Serial No"])    : e.data?.serialNo,
                size:         row["Size (DN)"]    !== undefined ? String(row["Size (DN)"])    : e.data?.size,
                spec:         row["Schedule"]     !== undefined ? String(row["Schedule"])     : e.data?.spec,
                fluidPrimary: row["Fluid Primary"]!== undefined ? String(row["Fluid Primary"]): e.data?.fluidPrimary,
                fluidSub:     row["Fluid Sub"]    !== undefined ? String(row["Fluid Sub"])    : e.data?.fluidSub,
                lineText:     row["Line Text"]    !== undefined ? String(row["Line Text"])    : e.data?.lineText,
              }
            };
          });
          setEdges(updatedEdges);
        }

        // ── Requirements 시트 ──
        const wsR = wb.Sheets["Requirements"];
        if (wsR) {
          const rows = XLSX.utils.sheet_to_json(wsR);
          // nodeId별로 그룹핑
          const reqMap = {};
          rows.forEach(r => {
            if (!r["Node ID"] || !r["Requirement"]) return;
            if (!reqMap[r["Node ID"]]) reqMap[r["Node ID"]] = [];
            reqMap[r["Node ID"]].push({
              id:   r["Req ID"] || Date.now() + Math.random(),
              text: String(r["Requirement"] || ""),
              who:  String(r["Stakeholder"] || ""),
              date: String(r["Date"] || ""),
            });
          });
          setNodes(ns => ns.map(n => {
            if (!reqMap[n.id]) return n;
            return { ...n, data:{ ...n.data, requirements: reqMap[n.id] } };
          }));
        }

        resolve("Excel import 완료");
      } catch(err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
};

// ─────────────────────────────────────────────────────────────
// GLOBAL CSS
// ─────────────────────────────────────────────────────────────
const GLOBAL_CSS = `
  .react-flow__handle {
    opacity: 0 !important;
    transition: opacity 0.15s ease, transform 0.15s ease !important;
  }
  .react-flow__node:hover .react-flow__handle { opacity: 1 !important; }
  .react-flow__node.selected .react-flow__handle { opacity: 1 !important; }
  .react-flow__handle:hover { opacity: 1 !important; transform: scale(1.4) !important; }
  .react-flow__handle.connecting { opacity: 1 !important; }
  /* 범위 선택 박스 스타일 */
  .react-flow__selection {
    background: rgba(59,130,246,0.08) !important;
    border: 1.5px dashed #3b82f6 !important;
    border-radius: 4px !important;
  }
  /* 선택된 노드 강조 */
  .react-flow__node.selected > div {
    box-shadow: 0 0 0 2px #3b82f6 !important;
  }
  .mbse-label-input {
    background: rgba(255,255,255,0.95);
    border: 1.5px solid #3b82f6;
    border-radius: 4px;
    padding: 2px 6px;
    font-size: 11px;
    font-weight: 700;
    text-align: center;
    outline: none;
    width: 100%;
    box-sizing: border-box;
  }
`;

// ─────────────────────────────────────────────────────────────
// CONSTANTS
// ─────────────────────────────────────────────────────────────
const AREA_TYPES = ["Plant","System","Package","Item"];
const AREA_COLORS = {
  Plant:    { bg:"rgba(219,234,254,0.45)", border:"#93c5fd", label:"#1d4ed8" },
  System:   { bg:"rgba(220,252,231,0.45)", border:"#86efac", label:"#15803d" },
  Package:  { bg:"rgba(254,249,195,0.45)", border:"#fde047", label:"#a16207" },
  Item:     { bg:"rgba(243,232,255,0.45)", border:"#d8b4fe", label:"#7e22ce" },
};
const EQUIPMENT_LIST = [
  "Tank","Pump","Pond","Heat Exchanger","Filter","Hopper","Decanter",
  "Cooling Tower","Clarifier","Classifier","Feed Box","Chemical Dosing",
  "Scrubber","Bag Filter","Reactor","Feed Bin","Gas Duct","Steel Structure",
];

// ── Connection 타입 ──────────────────────────────────────────
// Piping / Duct / Brench : 기존
// Process Gas            : 굵은 보라색 선, 텍스트 라벨 지원
// Material               : 굵은 갈색 선, 텍스트 라벨 지원
const CONNECTION_LIST = ["Piping","Duct","Brench","Process Gas","Material"];
const CONVEYOR_LIST   = ["Conveyor"];

// 라인 스타일 정의
const LINE_STYLE = {
  Piping:      { color:"#94a3b8", sw:2,   dash:"none"  },
  Duct:        { color:"#475569", sw:4,   dash:"none"  },
  Brench:      { color:"#94a3b8", sw:1.5, dash:"none"  },
  "Process Gas":{ color:"#7c3aed", sw:5,  dash:"none"  },
  Material:    { color:"#92400e", sw:5,   dash:"none"  },
  Conveyor:    { color:"#78350f", sw:2,   dash:"7,3"   },
};

const INSTRUMENT_CATS = { Flow:[], Pressure:[], Temperature:[], Level:[] };
const INSTR_TYPES     = ["Transmitter","Switch","Gauge"];
const FLUID_PRIMARY   = ["Water","Slurry Water","Gas","Process Gas","Steam","Chemical"];
const FLUID_SECONDARY = {
  Water:          ["FW","PW","WM","GW","SW","WW","MCWS","MCR","PWS","PWR","PCC"],
  "Slurry Water": ["PWR","PCC","SL","TW"],
  Gas:            ["NI","NIH","CA","IA","NG","OX"],
  "Process Gas":  ["H2","FG","LG","VG","RG","RD","OG","EOG"],
  Steam:          ["ST","CS"],
  Chemical:       ["BIO","CI","COA","DIS","FL","SI"],
};
const FLUID_COLORS = {
  FW:"#2563eb",PW:"#3b82f6",WM:"#60a5fa",GW:"#22d3ee",SW:"#0ea5e9",
  WW:"#64748b",MCWS:"#06b6d4",MCR:"#0891b2",PWS:"#818cf8",PWR:"#6366f1",
  PCC:"#a78bfa",SL:"#92400e",TW:"#b45309",NI:"#7c3aed",NIH:"#6d28d9",
  CA:"#16a34a",IA:"#15803d",NG:"#d97706",OX:"#dc2626",H2:"#9333ea",
  FG:"#c2410c",LG:"#b45309",VG:"#6b7280",RG:"#475569",RD:"#ef4444",
  OG:"#ea580c",EOG:"#f97316",ST:"#94a3b8",CS:"#cbd5e1",
  BIO:"#65a30d",CI:"#0891b2",COA:"#78350f",DIS:"#059669",FL:"#dc2626",SI:"#7c3aed",
};
const getFluidColor = s => FLUID_COLORS[s] || "#94a3b8";

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
const uid = (p="n") => `${p}_${Date.now()}_${_idCnt++}`;

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
    {ICONS[type] || (<><rect x="3" y="3" width="18" height="18" rx="3" fill="none" stroke="currentColor" strokeWidth="1.5"/><text x="12" y="15" textAnchor="middle" fontSize="8" fill="currentColor" fontWeight="600">{type.substring(0,3).toUpperCase()}</text></>)}
  </svg>
);

const DIRS = [Position.Top, Position.Bottom, Position.Left, Position.Right];
const DIR_ID = { [Position.Top]:"top",[Position.Bottom]:"bottom",[Position.Left]:"left",[Position.Right]:"right" };

// ─────────────────────────────────────────────────────────────
// AREA IO UTILITY
// ─────────────────────────────────────────────────────────────
const computeAreaIO = (areaNode, allNodes, allEdges) => {
  const ax=areaNode.position.x, ay=areaNode.position.y;
  const aw=areaNode.style?.width||areaNode.width||260;
  const ah=areaNode.style?.height||areaNode.height||180;
  const insideIds=new Set(
    allNodes.filter(n=>{
      if(n.type==="area") return false;
      return n.position.x>=ax&&n.position.y>=ay&&n.position.x<=ax+aw&&n.position.y<=ay+ah;
    }).map(n=>n.id)
  );
  if(insideIds.size===0) return { inlets:[],outlets:[] };
  const inlets=[],outlets=[];
  allEdges.forEach(e=>{
    const si=insideIds.has(e.source),ti=insideIds.has(e.target);
    if(si===ti) return;
    const d=e.data||{},sub=d.fluidSub||d.lineType||"—",size=d.size?` ${d.size}`:"";
    if(ti) inlets.push(`${sub}${size}`);
    else   outlets.push(`${sub}${size}`);
  });
  return { inlets,outlets };
};

// ─────────────────────────────────────────────────────────────
// AREA NODE
// Handle 추가: 기본 상하좌우 4개, Inspector에서 추가 가능
// ─────────────────────────────────────────────────────────────
const AreaNode = memo(({ id, data, selected }) => {
  const c = AREA_COLORS[data.areaType] || AREA_COLORS.Plant;
  const [editing,setEditing] = useState(false);
  const [draft,setDraft]     = useState("");
  const inputRef = useRef(null);
  const inlets  = data.autoInlets  || [];
  const outlets = data.autoOutlets || [];
  const hasIO   = inlets.length>0 || outlets.length>0;

  // handles — Area도 Equipment처럼 상하좌우 기본 + 추가 가능
  const dirs   = data.handles || ["top","bottom","left","right"];
  const posMap = { top:Position.Top,bottom:Position.Bottom,left:Position.Left,right:Position.Right };
  const sideCount={},sideCursor={};
  dirs.forEach(d=>{ sideCount[d]=(sideCount[d]||0)+1; });
  const handleList=dirs.map((dir,i)=>{
    sideCursor[dir]=(sideCursor[dir]||0);
    const idx=sideCursor[dir]++,total=sideCount[dir];
    const pct=total===1?50:20+(idx/(total-1))*60;
    return { dir,pid:`${dir}_${i}`,pct };
  });
  const getHStyle=(dir,pct)=>{
    const base={ width:10,height:10,borderRadius:"50%",background:c.label,border:"2px solid #fff",zIndex:20 };
    if(dir==="top"||dir==="bottom") return { ...base,left:`${pct}%`,transform:"translateX(-50%)" };
    return { ...base,top:`${pct}%`,transform:"translateY(-50%)" };
  };

  const startEdit=useCallback(e=>{ e.stopPropagation(); setDraft(data.label||""); setEditing(true); setTimeout(()=>inputRef.current?.select(),30); },[data.label]);
  const commitEdit=useCallback(()=>{ setEditing(false); window.dispatchEvent(new CustomEvent("mbse:updatelabel",{ detail:{ id,label:draft } })); },[id,draft]);

  return (
    <div style={{
      background:c.bg, border:`2px dashed ${selected?"#f59e0b":c.border}`,
      borderRadius:10, width:"100%", height:"100%",
      position:"relative", boxSizing:"border-box", overflow:"hidden",
    }}>
      <NodeResizer minWidth={180} minHeight={120} isVisible={selected}/>

      {/* Handles — Area 경계에 표시 */}
      {handleList.map(({dir,pid,pct})=>(
        <Handle key={pid} type="source" position={posMap[dir]} id={pid} style={getHStyle(dir,pct)}/>
      ))}

      {/* Title bar */}
      <div onDoubleClick={startEdit} style={{
        position:"absolute",top:0,left:0,right:0,
        padding:"10px 14px 8px",
        background:selected?"rgba(245,158,11,0.12)":"rgba(255,255,255,0.55)",
        borderBottom:`1px dashed ${c.border}`,
        cursor:"text",userSelect:"none",
        display:"flex", alignItems:"center", justifyContent:"space-between",
      }}>
        <div style={{ flex:1, minWidth:0 }}>
          {editing ? (
            <input ref={inputRef} className="mbse-label-input" value={draft}
              onChange={e=>setDraft(e.target.value)} onBlur={commitEdit}
              onKeyDown={e=>{ if(e.key==="Enter") commitEdit(); if(e.key==="Escape") setEditing(false); }}
              onClick={e=>e.stopPropagation()} style={{ color:c.label }}/>
          ) : (
            <span style={{ fontSize:30,fontWeight:800,color:c.label,lineHeight:1.2 }}>
              [{data.areaType}]{data.label?` ${data.label}`:" (더블클릭으로 편집)"}
            </span>
          )}
        </div>
        {/* IN / OUT 토글 버튼 */}
        {hasIO && (
          <div style={{ display:"flex", gap:4, marginLeft:8, flexShrink:0 }} onClick={e=>e.stopPropagation()}>
            {inlets.length>0 && (
              <button
                onDoubleClick={e=>e.stopPropagation()}
                onClick={e=>{ e.stopPropagation(); window.dispatchEvent(new CustomEvent("mbse:updatelabel",{ detail:{ id, toggleIO: !data.showIO } })); }}
                style={{
                  background: data.showIO ? "#1d4ed8" : "rgba(255,255,255,0.8)",
                  color: data.showIO ? "#fff" : "#1d4ed8",
                  border:`1.5px solid #1d4ed8`,
                  borderRadius:4, padding:"2px 8px", cursor:"pointer",
                  fontSize:11, fontWeight:700, lineHeight:1.4,
                }}>
                IN {inlets.length}
              </button>
            )}
            {outlets.length>0 && (
              <button
                onDoubleClick={e=>e.stopPropagation()}
                onClick={e=>{ e.stopPropagation(); window.dispatchEvent(new CustomEvent("mbse:updatelabel",{ detail:{ id, toggleIO: !data.showIO } })); }}
                style={{
                  background: data.showIO ? "#dc2626" : "rgba(255,255,255,0.8)",
                  color: data.showIO ? "#fff" : "#dc2626",
                  border:`1.5px solid #dc2626`,
                  borderRadius:4, padding:"2px 8px", cursor:"pointer",
                  fontSize:11, fontWeight:700, lineHeight:1.4,
                }}>
                OUT {outlets.length}
              </button>
            )}
          </div>
        )}
      </div>

      {/* IO summary — 토글 버튼으로 제어 */}
      {hasIO && (
        <div style={{ position:"absolute",top:55,left:8,right:8,userSelect:"none" }}>
          {data.showIO && (
            <div style={{ display:"flex",gap:6,pointerEvents:"none" }}>
              {inlets.length>0 && (
                <div style={{ flex:1,background:"rgba(255,255,255,0.92)",border:`1px solid ${c.border}`,borderRadius:5,padding:"4px 6px",fontSize:9,lineHeight:1.6,boxShadow:"0 2px 6px rgba(0,0,0,0.08)" }}>
                  <div style={{ fontWeight:800,color:"#1d4ed8",marginBottom:1 }}>▶ INLET</div>
                  {inlets.map((s,i)=>(
                    <div key={i} style={{ color:"#1e293b",display:"flex",alignItems:"center",gap:3 }}>
                      <div style={{ width:6,height:6,borderRadius:"50%",background:getFluidColor(s.split(" ")[0]),flexShrink:0 }}/>{s}
                    </div>
                  ))}
                </div>
              )}
              {outlets.length>0 && (
                <div style={{ flex:1,background:"rgba(255,255,255,0.92)",border:`1px solid ${c.border}`,borderRadius:5,padding:"4px 6px",fontSize:9,lineHeight:1.6,boxShadow:"0 2px 6px rgba(0,0,0,0.08)" }}>
                  <div style={{ fontWeight:800,color:"#dc2626",marginBottom:1 }}>◀ OUTLET</div>
                  {outlets.map((s,i)=>(
                    <div key={i} style={{ color:"#1e293b",display:"flex",alignItems:"center",gap:3 }}>
                      <div style={{ width:6,height:6,borderRadius:"50%",background:getFluidColor(s.split(" ")[0]),flexShrink:0 }}/>{s}
                    </div>
                  ))}
                </div>
              )}
            </div>
          )}
        </div>
      )}

      {data.summary && (
        <div style={{ position:"absolute",bottom:6,left:8,right:8,fontSize:9,color:c.label,lineHeight:1.4,userSelect:"none",pointerEvents:"none",whiteSpace:"pre-wrap",opacity:0.85 }}>
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
  const posMap = { top:Position.Top,bottom:Position.Bottom,left:Position.Left,right:Position.Right };
  const [editing,setEditing]=useState(false);
  const [draft,setDraft]=useState("");
  const inputRef=useRef(null);

  const sideCount={},sideCursor={};
  dirs.forEach(d=>{ sideCount[d]=(sideCount[d]||0)+1; });
  const handleList=dirs.map((dir,i)=>{
    sideCursor[dir]=(sideCursor[dir]||0);
    const idx=sideCursor[dir]++,total=sideCount[dir];
    const pct=total===1?50:20+(idx/(total-1))*60;
    return { dir,pid:`${dir}_${i}`,pct };
  });
  const getStyle=(dir,pct)=>{
    const base={ width:9,height:9,borderRadius:"50%",background:"#3b82f6",border:"2px solid #fff",zIndex:10 };
    if(dir==="top"||dir==="bottom") return { ...base,left:`${pct}%`,transform:"translateX(-50%)" };
    return { ...base,top:`${pct}%`,transform:"translateY(-50%)" };
  };
  const displayName=data.itemNo||data.equipType;
  const startEdit=useCallback(e=>{ e.stopPropagation(); setDraft(data.itemNo||""); setEditing(true); setTimeout(()=>inputRef.current?.select(),30); },[data.itemNo]);
  const commitEdit=useCallback(()=>{ setEditing(false); window.dispatchEvent(new CustomEvent("mbse:updatelabel",{ detail:{ id,itemNo:draft } })); },[id,draft]);

  return (
    <div style={{
      background:selected?"#eff6ff":"#ffffff",
      border:`${selected?2:1.5}px solid ${selected?"#3b82f6":"#cbd5e1"}`,
      borderRadius:8,minWidth:110,minHeight:64,
      display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",
      padding:"8px 12px",fontSize:11,cursor:"default",
      position:"relative",userSelect:"none",boxSizing:"border-box",
    }}>
      <NodeResizer minWidth={80} minHeight={50} isVisible={selected} handleStyle={{ width:8,height:8 }}/>
      {handleList.map(({dir,pid,pct})=>(
        <Handle key={pid} type="source" position={posMap[dir]} id={pid} style={getStyle(dir,pct)}/>
      ))}
      <div style={{ color:"#3b82f6",marginBottom:4 }}><EquipSVG type={data.equipType} size={26}/></div>
      {editing ? (
        <input ref={inputRef} className="mbse-label-input" value={draft}
          onChange={e=>setDraft(e.target.value)} onBlur={commitEdit}
          onKeyDown={e=>{ if(e.key==="Enter") commitEdit(); if(e.key==="Escape") setEditing(false); }}
          onClick={e=>e.stopPropagation()} placeholder="Item No"/>
      ) : (
        <div onDoubleClick={startEdit}
          style={{ fontWeight:700,color:"#0f172a",fontSize:11,textAlign:"center",lineHeight:1.3,cursor:"text",padding:"1px 4px",borderRadius:3,minWidth:60 }}
          title="더블클릭으로 Item No 편집">
          {displayName}
        </div>
      )}
      {data.label && data.label!==displayName && (
        <div style={{ color:"#64748b",fontSize:10,textAlign:"center" }}>{data.label}</div>
      )}
    </div>
  );
});

// ─────────────────────────────────────────────────────────────
// BRENCH NODE
// ─────────────────────────────────────────────────────────────
const BrenchNode = memo(({ data, selected }) => (
  <div style={{ width:18,height:18,borderRadius:"50%",background:selected?"#f59e0b":"#334155",border:`2px solid ${selected?"#b45309":"#1e293b"}`,position:"relative" }}>
    {DIRS.map(pos=>(
      <Handle key={pos} type="source" position={pos} id={DIR_ID[pos]}
        style={{ width:7,height:7,borderRadius:"50%",background:"#94a3b8",border:"1px solid #fff" }}/>
    ))}
  </div>
));

// ─────────────────────────────────────────────────────────────
// INSTRUMENT NODE
// ─────────────────────────────────────────────────────────────
const InstrumentNode = memo(({ data, selected }) => (
  <div style={{ width:54,height:54,borderRadius:"50%",background:selected?"#faf5ff":"#fafafa",border:`${selected?2:1.5}px solid ${selected?"#a855f7":"#c4b5fd"}`,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",fontSize:10,position:"relative",userSelect:"none" }}>
    {DIRS.map(pos=>(
      <Handle key={pos} type="source" position={pos} id={DIR_ID[pos]}
        style={{ width:7,height:7,background:"#a855f7",border:"1px solid #fff" }}/>
    ))}
    <div style={{ fontWeight:800,color:"#7c3aed",fontSize:12,lineHeight:1 }}>
      {(data.instrCategory||"?").charAt(0)}{(data.instrType||"?").charAt(0)}
    </div>
    <div style={{ color:"#94a3b8",fontSize:9 }}>{data.itemNo||""}</div>
  </div>
));

// ─────────────────────────────────────────────────────────────
// PIPE EDGE
// Process Gas / Material: 굵은 선 + 중앙 텍스트 라벨 (인라인 편집 포함)
// ─────────────────────────────────────────────────────────────
const PipeEdge = ({
  id,sourceX,sourceY,targetX,targetY,
  sourcePosition,targetPosition,data,selected,
}) => {
  const [edgePath] = getSmoothStepPath({ sourceX,sourceY,sourcePosition,targetX,targetY,targetPosition,borderRadius:4 });

  const lt = data?.lineType || "Piping";
  const ls = LINE_STYLE[lt] || LINE_STYLE.Piping;

  // Fluid 색상 우선, 없으면 lineType 기본색
  const baseColor = data?.fluidSub ? getFluidColor(data.fluidSub) : ls.color;
  const stroke    = selected ? "#f59e0b" : baseColor;
  const sw        = ls.sw;
  const mkId      = `mk_${id}`;

  // Process Gas / Material 은 lineText 라벨 표시
  const isSpecial = lt==="Process Gas" || lt==="Material";
  const labelText = data?.lineText || (isSpecial ? lt : null);

  // 라인 라벨: Fluid-SizeA 형식 (예: FW-200A)
  const fluidLabel = data?.fluidSub || "";
  const sizeLabel  = data?.sizeNum ? `${data.sizeNum}A` : (data?.size || "");
  const pipingLabel = [fluidLabel, sizeLabel].filter(Boolean).join("-");

  const showLabel = isSpecial ? labelText : pipingLabel;

  const mx = (sourceX+targetX)/2;
  const my = (sourceY+targetY)/2;

  return (
    <g>
      <defs>
        <marker id={mkId} markerWidth="5" markerHeight="4" refX="4.5" refY="2" orient="auto" markerUnits="strokeWidth">
          <polygon points="0 0, 5 2, 0 4" fill={stroke}/>
        </marker>
      </defs>
      {/* hit area */}
      <path d={edgePath} fill="none" stroke="transparent" strokeWidth={16} style={{ cursor:"pointer" }}/>
      {/* visible line */}
      <path d={edgePath} fill="none" stroke={stroke}
        strokeWidth={selected?sw+0.5:sw}
        strokeDasharray={ls.dash}
        markerEnd={`url(#${mkId})`}
        style={{ pointerEvents:"none" }}/>
      {/* label */}
      {showLabel && (
        <EdgeLabelRenderer>
          <div style={{
            position:"absolute",
            transform:`translate(-50%,-50%) translate(${mx}px,${my}px)`,
            fontSize:10, fontWeight:isSpecial?700:600,
            background:"rgba(255,255,255,0.92)",
            padding:"1px 6px", borderRadius:4,
            border:`1.5px solid ${baseColor}`,
            color:baseColor,
            pointerEvents:"none", whiteSpace:"nowrap",
            boxShadow:"0 1px 3px rgba(0,0,0,0.08)",
          }}>
            {showLabel}
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
// SIDEBAR
// ─────────────────────────────────────────────────────────────
const Sidebar = memo(({ onDragStart }) => {
  const [open,setOpen]=useState({ Area:true,Equipment:true,Connection:false,Instrument:false });
  const tog=k=>setOpen(p=>({...p,[k]:!p[k]}));
  const iS=(color="#475569")=>({ padding:"5px 14px 5px 20px",cursor:"grab",borderBottom:"1px solid #f1f5f9",color,fontSize:11,userSelect:"none",display:"flex",alignItems:"center",gap:6 });
  const hov={ onMouseEnter:e=>{e.currentTarget.style.background="#e0f2fe";},onMouseLeave:e=>{e.currentTarget.style.background="";} };
  const Sec=({title,cat,children})=>(
    <div>
      <div onClick={()=>tog(cat)} style={{ padding:"7px 14px",fontWeight:600,color:"#334155",cursor:"pointer",background:"#f1f5f9",borderBottom:"1px solid #e2e8f0",display:"flex",justifyContent:"space-between",userSelect:"none",fontSize:12 }}>
        <span>{title}</span><span style={{ fontSize:10 }}>{open[cat]?"▲":"▼"}</span>
      </div>
      {open[cat]&&children}
    </div>
  );

  // 라인 미리보기 색/굵기
  const linePreview=(lt)=>{
    const ls=LINE_STYLE[lt]||LINE_STYLE.Piping;
    return <span style={{ display:"inline-block",width:28,height:ls.sw>3?ls.sw:2,background:ls.color,borderRadius:2,flexShrink:0,verticalAlign:"middle",marginRight:4 }}/>;
  };

  return (
    <div style={{ width:195,background:"#f8fafc",borderRight:"1px solid #e2e8f0",overflowY:"auto",flexShrink:0 }}>
      <div style={{ padding:"10px 14px 7px",fontWeight:800,fontSize:13,color:"#0f172a",borderBottom:"1px solid #e2e8f0",letterSpacing:0.5 }}>Catalog</div>
      <Sec title="Area" cat="Area">
        {AREA_TYPES.map(at=>{ const c=AREA_COLORS[at]||{ bg:"rgba(219,234,254,0.45)", border:"#93c5fd", label:"#1d4ed8" }; return (
          <div key={at} draggable onDragStart={e=>onDragStart(e,"area",at)} style={{ ...iS(c?.label||"#1d4ed8"),background:c?.bg||"transparent" }} {...hov}>▭ {at}</div>
        );})}
      </Sec>
      <Sec title="Equipment" cat="Equipment">
        {EQUIPMENT_LIST.map(eq=>(
          <div key={eq} draggable onDragStart={e=>onDragStart(e,"equipment",eq)} style={iS()} {...hov}><EquipSVG type={eq} size={14}/>{eq}</div>
        ))}
      </Sec>
      <Sec title="Connection" cat="Connection">
        {CONNECTION_LIST.map(c=>(
          <div key={c} draggable onDragStart={e=>onDragStart(e,"connection",c)}
            style={{ ...iS(LINE_STYLE[c]?.color||"#1d4ed8") }} {...hov}>
            {linePreview(c)}{c}
          </div>
        ))}
        {CONVEYOR_LIST.map(c=>(
          <div key={c} draggable onDragStart={e=>onDragStart(e,"conveyor",c)} style={iS("#78350f")} {...hov}>
            {linePreview("Conveyor")}╌╌ {c}
          </div>
        ))}
      </Sec>
      <Sec title="Instrument" cat="Instrument">
        {Object.keys(INSTRUMENT_CATS).map(cat=>
          INSTR_TYPES.map(t=>(
            <div key={`${cat}-${t}`} draggable onDragStart={e=>onDragStart(e,"instrument",`${cat}|${t}`)} style={iS("#7c3aed")} {...hov}>⊙ {cat} {t}</div>
          ))
        )}
      </Sec>
    </div>
  );
});

// ─────────────────────────────────────────────────────────────
// INSPECTOR
// ─────────────────────────────────────────────────────────────
const Inspector = memo(({ sel,nodes,edges,onUpdateNode,onUpdateEdge,onDeleteSel,onAddHandle }) => {
  const [reqText,setReqText]=useState(""), [reqWho,setReqWho]=useState(""), [reqDate,setReqDate]=useState(new Date().toISOString().slice(0,10));
  const [tab,setTab]=useState("spec");
  const L={fontSize:11,color:"#64748b",marginBottom:2,display:"block",fontWeight:500};
  const I={width:"100%",padding:"4px 7px",border:"1px solid #e2e8f0",borderRadius:4,fontSize:12,boxSizing:"border-box",marginBottom:7};
  const S={...I,background:"#fff"};

  if(!sel) return (
    <div style={{ width:270,background:"#fff",borderLeft:"1px solid #e2e8f0",padding:16,color:"#94a3b8",fontSize:12,flexShrink:0 }}>
      <div style={{ fontWeight:800,fontSize:13,color:"#0f172a",marginBottom:10 }}>Inspector</div>
      노드 또는 엣지를 선택하세요.
      <div style={{ marginTop:16,padding:10,background:"#f8fafc",borderRadius:6,fontSize:11,color:"#64748b",lineHeight:1.8 }}>
        <div style={{ fontWeight:700,color:"#334155",marginBottom:4 }}>단축키 안내</div>
        <div>노드 선택 후 <b>Ctrl+C</b> → 복사</div>
        <div><b>Ctrl+V</b> → 붙여넣기</div>
        <div>Area/Equipment <b>더블클릭</b> → 이름 편집</div>
        <div><b>Del</b> → 삭제</div>
      </div>
    </div>
  );

  const isNode=!!sel.position, d=sel.data||{};
  const upN=(k,v)=>onUpdateNode(sel.id,{...d,[k]:v});
  const upE=(k,v)=>onUpdateEdge(sel.id,{...d,[k]:v});
  const addReq=()=>{
    if(!reqText.trim()) return;
    const reqs=[...(d.requirements||[]),{id:Date.now(),text:reqText,who:reqWho,date:reqDate}];
    onUpdateNode(sel.id,{...d,requirements:reqs}); setReqText(""); setReqWho("");
  };
  const connEdges=edges.filter(e=>e.source===sel.id||e.target===sel.id);
  const ifaceList=connEdges.map(e=>{
    const other=e.source===sel.id?nodes.find(n=>n.id===e.target):nodes.find(n=>n.id===e.source);
    return { edgeId:e.id,dir:e.source===sel.id?"→":"←",otherLabel:other?.data?.itemNo||other?.data?.label||other?.id||"?",...e.data };
  });
  const TB=({name,label})=>(
    <button onClick={()=>setTab(name)} style={{ flex:1,padding:"5px 0",border:"none",cursor:"pointer",fontSize:11,background:tab===name?"#eff6ff":"#f8fafc",color:tab===name?"#1d4ed8":"#64748b",fontWeight:tab===name?700:400,borderBottom:tab===name?"2px solid #3b82f6":"2px solid transparent" }}>{label}</button>
  );

  const isSpecialLine = !isNode && (d.lineType==="Process Gas"||d.lineType==="Material");

  return (
    <div style={{ width:270,background:"#fff",borderLeft:"1px solid #e2e8f0",overflowY:"auto",flexShrink:0,display:"flex",flexDirection:"column" }}>
      <div style={{ padding:"10px 12px 7px",borderBottom:"1px solid #e2e8f0",display:"flex",justifyContent:"space-between",alignItems:"center",flexShrink:0 }}>
        <span style={{ fontWeight:800,fontSize:13,color:"#0f172a" }}>Inspector</span>
        <button onClick={onDeleteSel} style={{ background:"#fee2e2",color:"#dc2626",border:"none",borderRadius:4,padding:"2px 9px",cursor:"pointer",fontSize:11 }}>Delete</button>
      </div>
      {isNode && (
        <div style={{ display:"flex",borderBottom:"1px solid #e2e8f0",flexShrink:0 }}>
          <TB name="spec" label="Spec"/><TB name="req" label="Req."/><TB name="iface" label="I/F"/>
          {sel.type==="area"&&<TB name="io" label="IO"/>}
        </div>
      )}
      <div style={{ padding:"10px 12px",overflowY:"auto",flex:1 }}>

        {(tab==="spec"||!isNode) && (
          <>
            {/* AREA */}
            {isNode&&sel.type==="area"&&(
              <>
                <label style={L}>Area Type</label>
                <select style={S} value={d.areaType||"Plant"} onChange={e=>upN("areaType",e.target.value)}>
                  {AREA_TYPES.map(a=><option key={a}>{a}</option>)}
                </select>
                <label style={L}>Label (또는 더블클릭)</label>
                <input style={I} value={d.label||""} onChange={e=>upN("label",e.target.value)}/>
                <label style={L}>Notes</label>
                <textarea style={{ ...I,height:70,resize:"vertical" }} value={d.summary||""} onChange={e=>upN("summary",e.target.value)}/>
                <div style={{ fontWeight:600,fontSize:11,color:"#334155",margin:"6px 0 5px",borderTop:"1px solid #f1f5f9",paddingTop:5 }}>Port Management</div>
                <div style={{ display:"flex",flexWrap:"wrap",gap:5,marginBottom:6 }}>
                  {["top","bottom","left","right"].map(dir=>(
                    <button key={dir} onClick={()=>onAddHandle(sel.id,dir)} style={{ background:"#eff6ff",border:"1px solid #bfdbfe",color:"#1d4ed8",borderRadius:4,padding:"2px 8px",cursor:"pointer",fontSize:11 }}>+ {dir}</button>
                  ))}
                </div>
                <div style={{ fontSize:10,color:"#94a3b8" }}>Ports: {(d.handles||["top","bottom","left","right"]).join(", ")}</div>
              </>
            )}
            {/* EQUIPMENT */}
            {isNode&&sel.type==="equipment"&&(
              <>
                <label style={L}>Item No (또는 더블클릭)</label>
                <input style={I} value={d.itemNo||""} onChange={e=>upN("itemNo",e.target.value)} placeholder="e.g. T-101"/>
                <label style={L}>Description</label>
                <input style={I} value={d.label||""} onChange={e=>upN("label",e.target.value)}/>
                <div style={{ fontWeight:600,fontSize:11,color:"#334155",margin:"6px 0 5px",borderTop:"1px solid #f1f5f9",paddingTop:5 }}>Engineering Spec</div>
                {[["material","Material"],["capacity","Capacity"],["designP","Design Pressure"],["designT","Design Temp"]].map(([k,l])=>(
                  <React.Fragment key={k}><label style={L}>{l}</label><input style={I} value={d[k]||""} onChange={e=>upN(k,e.target.value)}/></React.Fragment>
                ))}
                <div style={{ fontWeight:600,fontSize:11,color:"#334155",margin:"6px 0 5px",borderTop:"1px solid #f1f5f9",paddingTop:5 }}>Port Management</div>
                <div style={{ display:"flex",flexWrap:"wrap",gap:5,marginBottom:6 }}>
                  {["top","bottom","left","right"].map(dir=>(
                    <button key={dir} onClick={()=>onAddHandle(sel.id,dir)} style={{ background:"#eff6ff",border:"1px solid #bfdbfe",color:"#1d4ed8",borderRadius:4,padding:"2px 8px",cursor:"pointer",fontSize:11 }}>+ {dir}</button>
                  ))}
                </div>
                <div style={{ fontSize:10,color:"#94a3b8" }}>Ports: {(d.handles||[]).join(", ")}</div>
              </>
            )}
            {/* INSTRUMENT */}
            {isNode&&sel.type==="instrument"&&(
              <>
                <label style={L}>Item No</label>
                <input style={I} value={d.itemNo||""} onChange={e=>upN("itemNo",e.target.value)}/>
                <label style={L}>Category</label>
                <select style={S} value={d.instrCategory||""} onChange={e=>upN("instrCategory",e.target.value)}>
                  <option value="">Select</option>{Object.keys(INSTRUMENT_CATS).map(c=><option key={c}>{c}</option>)}
                </select>
                <label style={L}>Type</label>
                <select style={S} value={d.instrType||""} onChange={e=>upN("instrType",e.target.value)}>
                  <option value="">Select</option>{INSTR_TYPES.map(t=><option key={t}>{t}</option>)}
                </select>
              </>
            )}
            {isNode&&sel.type==="brench"&&<div style={{ fontSize:12,color:"#64748b" }}>Brench (Junction) node.<br/>배관 분기점으로 사용합니다.</div>}

            {/* EDGE */}
            {!isNode&&(
              <>
                <label style={L}>Line Type</label>
                <select style={S} value={d.lineType||"Piping"} onChange={e=>upE("lineType",e.target.value)}>
                  {[...CONNECTION_LIST,...CONVEYOR_LIST].map(c=><option key={c}>{c}</option>)}
                </select>

                {/* Process Gas / Material 전용: 라인 텍스트 */}
                {isSpecialLine&&(
                  <>
                    <div style={{ fontWeight:600,fontSize:11,color:"#334155",margin:"6px 0 5px",borderTop:"1px solid #f1f5f9",paddingTop:5 }}>
                      Line Label
                    </div>
                    <label style={L}>텍스트 (라인 위에 표시)</label>
                    <input style={I} value={d.lineText||""} onChange={e=>upE("lineText",e.target.value)} placeholder={d.lineType}/>
                    <div style={{ display:"flex",alignItems:"center",gap:6,marginBottom:8 }}>
                      <div style={{ width:40,height:LINE_STYLE[d.lineType]?.sw||5,background:LINE_STYLE[d.lineType]?.color,borderRadius:2 }}/>
                      <span style={{ fontSize:11,color:"#64748b" }}>{d.lineType} 라인</span>
                    </div>
                  </>
                )}

                {/* Piping 전용: 기존 스펙 */}
                {(d.lineType==="Piping"||!d.lineType)&&(
                  <>
                    <label style={L}>Serial No</label>
                    <input style={I} value={d.serialNo||""} onChange={e=>upE("serialNo",e.target.value)} placeholder="e.g. L-101"/>

                    <label style={L}>Size (A)</label>
                    <div style={{ display:"flex", alignItems:"center", gap:6, marginBottom:7 }}>
                      <input
                        style={{ ...I, marginBottom:0, flex:1 }}
                        value={d.sizeNum||""}
                        onChange={e=>{
                          const v=e.target.value.replace(/[^0-9]/g,"");
                          onUpdateEdge(sel.id,{ ...d, sizeNum:v, size: v ? `${v}A` : "" });
                        }}
                        placeholder="숫자만 입력 (e.g. 200)"
                        type="number" min="0"
                      />
                      <span style={{ fontSize:13, fontWeight:700, color:"#334155", flexShrink:0 }}>
                        {d.sizeNum ? `${d.sizeNum}A` : "— A"}
                      </span>
                    </div>

                    <label style={L}>Schedule</label>
                    <select style={S} value={d.spec||""} onChange={e=>{
                      if(e.target.value==="직접입력") return;
                      upE("spec",e.target.value);
                    }}>
                      <option value="">Select</option>
                      <option value="SCH20">SCH20</option>
                      <option value="SCH40">SCH40</option>
                      <option value="SCH80">SCH80</option>
                      <option value="직접입력">직접입력...</option>
                    </select>
                    {(d.spec && !["SCH20","SCH40","SCH80",""].includes(d.spec)) && (
                      <input style={I} value={d.spec||""} onChange={e=>upE("spec",e.target.value)} placeholder="직접 입력"/>
                    )}

                    <div style={{ fontWeight:600,fontSize:11,color:"#334155",margin:"6px 0 5px",borderTop:"1px solid #f1f5f9",paddingTop:5 }}>Fluid</div>
                    <label style={L}>Primary</label>
                    <select style={S} value={d.fluidPrimary||""} onChange={e=>onUpdateEdge(sel.id,{...d,fluidPrimary:e.target.value,fluidSub:""})}>
                      <option value="">Select</option>{FLUID_PRIMARY.map(f=><option key={f}>{f}</option>)}
                    </select>
                    {d.fluidPrimary&&FLUID_SECONDARY[d.fluidPrimary]&&(
                      <>
                        <label style={L}>Sub-type</label>
                        <select style={S} value={d.fluidSub||""} onChange={e=>upE("fluidSub",e.target.value)}>
                          <option value="">Select</option>{FLUID_SECONDARY[d.fluidPrimary].map(f=><option key={f}>{f}</option>)}
                        </select>
                      </>
                    )}
                    {d.fluidSub&&(
                      <div style={{ display:"flex",alignItems:"center",gap:8,marginTop:4,marginBottom:8 }}>
                        <div style={{ width:28,height:8,borderRadius:4,background:getFluidColor(d.fluidSub) }}/>
                        <span style={{ fontSize:11,color:"#64748b",fontWeight:600 }}>
                          {d.fluidSub}{d.sizeNum?`-${d.sizeNum}A`:""}
                        </span>
                      </div>
                    )}
                  </>
                )}
              </>
            )}
          </>
        )}

        {tab==="req"&&isNode&&(
          <>
            {(d.requirements||[]).map(r=>(
              <div key={r.id} style={{ background:"#fffbeb",border:"1px solid #fde68a",borderRadius:6,padding:"6px 8px",marginBottom:6,fontSize:11 }}>
                <div style={{ display:"flex",justifyContent:"space-between",marginBottom:2 }}>
                  <span style={{ fontWeight:700,color:"#92400e" }}>{r.who||"—"}</span>
                  <span style={{ color:"#a16207",fontSize:10 }}>{r.date}</span>
                </div>
                <div style={{ color:"#1c1917",lineHeight:1.4 }}>{r.text}</div>
                <button onClick={()=>{ const rs=(d.requirements||[]).filter(x=>x.id!==r.id); onUpdateNode(sel.id,{...d,requirements:rs}); }}
                  style={{ background:"none",border:"none",color:"#dc2626",cursor:"pointer",fontSize:10,padding:"2px 0",marginTop:2 }}>× remove</button>
              </div>
            ))}
            {!(d.requirements||[]).length&&<div style={{ color:"#94a3b8",fontSize:11,marginBottom:10 }}>등록된 요구사항 없음</div>}
            <div style={{ borderTop:"1px solid #f1f5f9",paddingTop:8 }}>
              <textarea style={{ ...I,height:56,resize:"vertical" }} placeholder="Requirement..." value={reqText} onChange={e=>setReqText(e.target.value)}/>
              <input style={I} placeholder="Stakeholder" value={reqWho} onChange={e=>setReqWho(e.target.value)}/>
              <input type="date" style={I} value={reqDate} onChange={e=>setReqDate(e.target.value)}/>
              <button onClick={addReq} style={{ width:"100%",padding:"6px",background:"#1d4ed8",color:"#fff",border:"none",borderRadius:5,cursor:"pointer",fontSize:12,fontWeight:700 }}>+ ADD Requirement</button>
            </div>
          </>
        )}

        {tab==="iface"&&isNode&&(
          <>
            <div style={{ fontWeight:600,fontSize:11,color:"#334155",marginBottom:6 }}>Connected Interfaces ({ifaceList.length})</div>
            {ifaceList.length===0&&<div style={{ color:"#94a3b8",fontSize:11 }}>연결된 인터페이스 없음</div>}
            {ifaceList.map(iface=>(
              <div key={iface.edgeId} style={{ border:"1px solid #e2e8f0",borderRadius:6,padding:"6px 8px",marginBottom:5,fontSize:11 }}>
                <div style={{ display:"flex",justifyContent:"space-between" }}>
                  <span style={{ fontWeight:700,color:"#1d4ed8" }}>{iface.dir} {iface.otherLabel}</span>
                  <span style={{ color:"#64748b" }}>{iface.lineType||"Piping"}</span>
                </div>
                {iface.fluidSub&&(
                  <div style={{ display:"flex",alignItems:"center",gap:5,marginTop:3 }}>
                    <div style={{ width:16,height:5,borderRadius:3,background:getFluidColor(iface.fluidSub) }}/>
                    <span style={{ color:"#64748b" }}>{iface.fluidSub} {iface.size||""} {iface.serialNo||""}</span>
                  </div>
                )}
                {iface.lineText&&<div style={{ color:"#64748b",marginTop:2 }}>Label: {iface.lineText}</div>}
              </div>
            ))}
          </>
        )}

        {tab==="io"&&isNode&&sel.type==="area"&&(
          <>
            <div style={{ fontWeight:700,fontSize:11,color:"#1d4ed8",marginBottom:8 }}>▶ INLET ({(d.autoInlets||[]).length})</div>
            {(d.autoInlets||[]).length===0?<div style={{ color:"#94a3b8",fontSize:11,marginBottom:10 }}>없음</div>
              :(d.autoInlets||[]).map((s,i)=>(
                <div key={i} style={{ display:"flex",alignItems:"center",gap:8,padding:"4px 8px",background:"#eff6ff",borderRadius:5,marginBottom:4,fontSize:11 }}>
                  <div style={{ width:10,height:10,borderRadius:"50%",background:getFluidColor(s.split(" ")[0]),flexShrink:0 }}/>
                  <span style={{ fontWeight:600,color:"#1e293b" }}>{s}</span>
                </div>
              ))
            }
            <div style={{ fontWeight:700,fontSize:11,color:"#dc2626",marginBottom:8,marginTop:10 }}>◀ OUTLET ({(d.autoOutlets||[]).length})</div>
            {(d.autoOutlets||[]).length===0?<div style={{ color:"#94a3b8",fontSize:11,marginBottom:10 }}>없음</div>
              :(d.autoOutlets||[]).map((s,i)=>(
                <div key={i} style={{ display:"flex",alignItems:"center",gap:8,padding:"4px 8px",background:"#fff1f2",borderRadius:5,marginBottom:4,fontSize:11 }}>
                  <div style={{ width:10,height:10,borderRadius:"50%",background:getFluidColor(s.split(" ")[0]),flexShrink:0 }}/>
                  <span style={{ fontWeight:600,color:"#1e293b" }}>{s}</span>
                </div>
              ))
            }
          </>
        )}
      </div>
    </div>
  );
});

// ─────────────────────────────────────────────────────────────
// CONNECTION MODAL — default 값 lineType 반영
// ─────────────────────────────────────────────────────────────
// ─────────────────────────────────────────────────────────────
// SAVE MEMO MODAL — 저장 시 수정자·수정내용 입력
// ─────────────────────────────────────────────────────────────
const SaveMemoModal = ({ nodes, edges, onConfirm, onCancel }) => {
  const [author,  setAuthor]  = useState("");
  const [summary, setSummary] = useState("");

  const inp = { width:"100%",padding:"7px 10px",border:"1px solid #e2e8f0",borderRadius:6,fontSize:13,boxSizing:"border-box",marginBottom:12,outline:"none" };

  // 변경 내용 자동 분석
  const autoDesc = `노드 ${nodes.length}개 · 연결 ${edges.length}개`;

  return (
    <div style={{ position:"fixed",inset:0,background:"rgba(0,0,0,0.5)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:9999 }}>
      <div style={{ background:"#fff",borderRadius:12,padding:28,minWidth:340,maxWidth:420,boxShadow:"0 12px 40px rgba(0,0,0,0.2)" }}>
        <div style={{ fontWeight:800,fontSize:16,marginBottom:6,color:"#0f172a" }}>💾 저장 메모</div>
        <div style={{ fontSize:11,color:"#64748b",marginBottom:18 }}>저장 시 기록할 내용을 입력해 주세요.</div>

        <label style={{ fontSize:12,color:"#334155",fontWeight:600,display:"block",marginBottom:4 }}>수정자 *</label>
        <input style={inp} value={author} onChange={e=>setAuthor(e.target.value)} placeholder="이름 또는 부서" autoFocus/>

        <label style={{ fontSize:12,color:"#334155",fontWeight:600,display:"block",marginBottom:4 }}>수정 내용 *</label>
        <textarea style={{ ...inp,height:80,resize:"vertical",fontFamily:"inherit" }}
          value={summary} onChange={e=>setSummary(e.target.value)}
          placeholder="어떤 내용을 수정했는지 간략히 입력하세요.&#10;예) Tank T-101 추가, FW 배관 연결"/>

        <div style={{ background:"#f8fafc",border:"1px solid #e2e8f0",borderRadius:6,padding:"8px 12px",marginBottom:16,fontSize:11,color:"#64748b" }}>
          <span style={{ fontWeight:600 }}>현재 상태: </span>{autoDesc}
        </div>

        <div style={{ display:"flex",gap:8 }}>
          <button
            onClick={()=>{ if(!author.trim()||!summary.trim()){ alert("수정자와 수정 내용을 입력해주세요."); return; } onConfirm(author.trim(),summary.trim()); }}
            style={{ flex:1,background:"#1d4ed8",color:"#fff",border:"none",borderRadius:7,padding:"10px",cursor:"pointer",fontWeight:700,fontSize:13 }}>
            저장
          </button>
          <button onClick={onCancel}
            style={{ flex:1,background:"#f1f5f9",color:"#334155",border:"1px solid #e2e8f0",borderRadius:7,padding:"10px",cursor:"pointer",fontSize:13 }}>
            취소
          </button>
        </div>
      </div>
    </div>
  );
};

// ─────────────────────────────────────────────────────────────
// LOG DETAIL MODAL — 로그 상세보기
// ─────────────────────────────────────────────────────────────
const LogDetailModal = ({ log, onClose }) => {
  if(!log) return null;
  return (
    <div style={{ position:"fixed",inset:0,background:"rgba(0,0,0,0.5)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:9999 }}>
      <div style={{ background:"#fff",borderRadius:12,padding:28,minWidth:380,maxWidth:500,boxShadow:"0 12px 40px rgba(0,0,0,0.2)",maxHeight:"80vh",overflowY:"auto" }}>
        <div style={{ display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16 }}>
          <div style={{ fontWeight:800,fontSize:15,color:"#0f172a" }}>📋 저장 이력 상세</div>
          <button onClick={onClose} style={{ background:"#f1f5f9",border:"none",borderRadius:5,padding:"3px 10px",cursor:"pointer",fontSize:12,color:"#64748b" }}>✕ 닫기</button>
        </div>

        {[
          ["📅 저장 일시", log.date],
          ["👤 수정자",     log.author],
          ["📝 수정 내용",  log.summary],
          ["📊 저장 시 상태", `노드 ${log.nodeCount}개 · 연결 ${log.edgeCount}개`],
          ["💾 파일명",     log.filename],
        ].map(([label,val])=>(
          <div key={label} style={{ marginBottom:14 }}>
            <div style={{ fontSize:11,color:"#64748b",fontWeight:600,marginBottom:3 }}>{label}</div>
            <div style={{ fontSize:13,color:"#1e293b",background:"#f8fafc",border:"1px solid #e2e8f0",borderRadius:6,padding:"8px 12px",lineHeight:1.6,whiteSpace:"pre-wrap" }}>
              {val||"—"}
            </div>
          </div>
        ))}
      </div>
    </div>
  );
};

const ConnModal = ({ onConfirm, onCancel, defaultType="Piping" }) => {
  const [lt,setLt]=useState(defaultType);
  useEffect(()=>setLt(defaultType),[defaultType]);
  return (
    <div style={{ position:"fixed",inset:0,background:"rgba(0,0,0,0.45)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:9999 }}>
      <div style={{ background:"#fff",borderRadius:10,padding:24,minWidth:260,boxShadow:"0 8px 30px rgba(0,0,0,0.18)" }}>
        <div style={{ fontWeight:800,fontSize:15,marginBottom:14,color:"#0f172a" }}>New Connection</div>
        <label style={{ fontSize:12,color:"#64748b",display:"block",marginBottom:4 }}>Line Type</label>
        <select value={lt} onChange={e=>setLt(e.target.value)} style={{ width:"100%",padding:"6px 8px",border:"1px solid #e2e8f0",borderRadius:6,fontSize:13,marginBottom:16 }}>
          {[...CONNECTION_LIST,...CONVEYOR_LIST].map(c=><option key={c}>{c}</option>)}
        </select>
        {/* 색상 미리보기 */}
        <div style={{ display:"flex",alignItems:"center",gap:8,marginBottom:16 }}>
          <div style={{ width:50,height:LINE_STYLE[lt]?.sw||2,background:LINE_STYLE[lt]?.color||"#94a3b8",borderRadius:2 }}/>
          <span style={{ fontSize:11,color:"#64748b" }}>{lt}</span>
        </div>
        <div style={{ display:"flex",gap:8 }}>
          <button onClick={()=>onConfirm(lt)} style={{ flex:1,background:"#1d4ed8",color:"#fff",border:"none",borderRadius:6,padding:"8px",cursor:"pointer",fontWeight:700 }}>Create</button>
          <button onClick={onCancel} style={{ flex:1,background:"#f1f5f9",color:"#334155",border:"1px solid #e2e8f0",borderRadius:6,padding:"8px",cursor:"pointer" }}>Cancel</button>
        </div>
      </div>
    </div>
  );
};

// ─────────────────────────────────────────────────────────────
// MAIN CANVAS
// ─────────────────────────────────────────────────────────────
const CanvasInner = () => {
  const wrapRef=useRef(null);
  const { screenToFlowPosition, fitView }=useReactFlow();
  const [nodes,setNodes,onNodesChange]=useNodesState([]);
  const [edges,setEdges,onEdgesChange]=useEdgesState([]);
  const [sel,setSel]       = useState(null);
  const [modal,setModal]   = useState(false);
  const [modalDefault,setModalDefault] = useState("Piping");
  const connRef     = useRef(null);
  const fileRef     = useRef(null);
  const xlsxRef     = useRef(null);
  const clipRef     = useRef(null);
  const [xlsxMsg, setXlsxMsg]     = useState("");
  const [saveMsg,  setSaveMsg]     = useState("");
  const [showLog,  setShowLog]     = useState(false);
  const [history,  setHistory]     = useState([]);     // 저장 이력 (Save 시에만 기록)
  const [showSaveModal, setShowSaveModal] = useState(false); // 저장 시 메모 모달
  const [logDetail, setLogDetail]  = useState(null);   // 상세보기 대상 log
  const autoSaveTimer = useRef(null);

  // ── Undo / Redo 스택 ──────────────────────────────────────
  const undoStack   = useRef([]);
  const redoStack   = useRef([]);
  const isUndoRedo  = useRef(false);
  const prevNodesRef = useRef(null);
  const prevEdgesRef = useRef(null);

  useEffect(()=>{
    if(isUndoRedo.current) return;
    if(prevNodesRef.current===null){ prevNodesRef.current=nodes; prevEdgesRef.current=edges; return; }
    const nChanged = JSON.stringify(prevNodesRef.current)!==JSON.stringify(nodes);
    const eChanged = JSON.stringify(prevEdgesRef.current)!==JSON.stringify(edges);
    if(nChanged||eChanged){
      undoStack.current.push({ nodes:prevNodesRef.current, edges:prevEdgesRef.current });
      if(undoStack.current.length>50) undoStack.current.shift();
      redoStack.current=[];
      prevNodesRef.current=nodes; prevEdgesRef.current=edges;
    }
  },[nodes,edges]); // eslint-disable-line

  const undo = useCallback(()=>{
    if(undoStack.current.length===0) return;
    const prev=undoStack.current.pop();
    redoStack.current.push({ nodes,edges });
    isUndoRedo.current=true;
    setNodes(prev.nodes); setEdges(prev.edges); setSel(null);
    setTimeout(()=>{ isUndoRedo.current=false; },50);
  },[nodes,edges,setNodes,setEdges]);

  const redo = useCallback(()=>{
    if(redoStack.current.length===0) return;
    const next=redoStack.current.pop();
    undoStack.current.push({ nodes,edges });
    isUndoRedo.current=true;
    setNodes(next.nodes); setEdges(next.edges); setSel(null);
    setTimeout(()=>{ isUndoRedo.current=false; },50);
  },[nodes,edges,setNodes,setEdges]);

  // ── 앱 시작 시 localStorage 복원 ─────────────────────────
  useEffect(()=>{
    try {
      const saved = localStorage.getItem("mbse_autosave");
      if(saved){
        const { nodes:n, edges:e, history:h } = JSON.parse(saved);
        if(n) setNodes(n);
        if(e) setEdges(e);
        if(h) setHistory(h);
        setSaveMsg("✅ 복원됨");
        setTimeout(()=>setSaveMsg(""),2000);
      }
    } catch(err){ console.warn("복원 실패",err); }
  },[]); // eslint-disable-line

  // ── 자동저장: 2초마다 localStorage (Log 기록 없이 데이터만) ─
  useEffect(()=>{
    if(nodes.length===0 && edges.length===0) return;
    if(autoSaveTimer.current) clearTimeout(autoSaveTimer.current);
    autoSaveTimer.current = setTimeout(()=>{
      try {
        localStorage.setItem("mbse_autosave", JSON.stringify({ nodes, edges, history }));
        setSaveMsg("💾 자동저장");
        setTimeout(()=>setSaveMsg(""),1200);
      } catch(err){ console.warn("자동저장 실패",err); }
    }, 2000);
    return()=>{ if(autoSaveTimer.current) clearTimeout(autoSaveTimer.current); };
  },[nodes,edges]); // eslint-disable-line

  // ── Save 버튼 클릭 → 메모 모달 오픈 ─────────────────────
  const onSave = () => setShowSaveModal(true);

  // ── 메모 입력 확인 → 파일 저장 + Log 기록 ────────────────
  const confirmSave = (author, summary) => {
    const now = new Date();
    const ts  = now.toISOString().replace(/[:.]/g,"-").slice(0,19);
    const filename = `MBSE_${ts}.json`;
    // 파일 다운로드
    const blob = new Blob([JSON.stringify({ nodes,edges },null,2)],{ type:"application/json" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob); a.download = filename; a.click();
    // Log 기록
    const dateStr = now.toLocaleString("ko-KR",{ year:"numeric",month:"2-digit",day:"2-digit",hour:"2-digit",minute:"2-digit" });
    const log = {
      id:       Date.now(),
      date:     dateStr,
      author,
      summary,
      filename,
      nodeCount: nodes.length,
      edgeCount: edges.length,
    };
    setHistory(h=>{
      const updated=[log,...h].slice(0,200);
      localStorage.setItem("mbse_autosave",JSON.stringify({ nodes,edges,history:updated }));
      return updated;
    });
    setShowSaveModal(false);
    setSaveMsg("✅ 저장 완료!");
    setTimeout(()=>setSaveMsg(""),2000);
  };

  // AREA IO 자동계산
  useEffect(()=>{
    const areaNodes=nodes.filter(n=>n.type==="area");
    if(areaNodes.length===0) return;
    let changed=false;
    const updated=nodes.map(n=>{
      if(n.type!=="area") return n;
      const {inlets,outlets}=computeAreaIO(n,nodes,edges);
      const same=JSON.stringify(n.data.autoInlets)===JSON.stringify(inlets)&&JSON.stringify(n.data.autoOutlets)===JSON.stringify(outlets);
      if(same) return n;
      changed=true;
      return { ...n,data:{ ...n.data,autoInlets:inlets,autoOutlets:outlets } };
    });
    if(changed) setNodes(updated);
  },[edges,nodes.map(n=>n.position.x+","+n.position.y).join("|")]); // eslint-disable-line

  // 인라인 편집 이벤트 수신
  useEffect(()=>{
    const fn=e=>{
      const {id,label,itemNo,toggleIO}=e.detail;
      setNodes(ns=>ns.map(n=>{
        if(n.id!==id) return n;
        if(label!==undefined)    return { ...n,data:{ ...n.data,label } };
        if(itemNo!==undefined)   return { ...n,data:{ ...n.data,itemNo } };
        if(toggleIO!==undefined) return { ...n,data:{ ...n.data,showIO:toggleIO } };
        return n;
      }));
    };
    window.addEventListener("mbse:updatelabel",fn);
    return()=>window.removeEventListener("mbse:updatelabel",fn);
  },[setNodes]);

  // DRAG
  const onDragStart=useCallback((e,cat,sub)=>{ e.dataTransfer.setData("mbse/cat",cat); e.dataTransfer.setData("mbse/sub",sub); e.dataTransfer.effectAllowed="move"; },[]);
  const onDragOver=useCallback(e=>{ e.preventDefault(); e.dataTransfer.dropEffect="move"; },[]);
  const onDrop=useCallback(e=>{
    e.preventDefault();
    const cat=e.dataTransfer.getData("mbse/cat"),sub=e.dataTransfer.getData("mbse/sub");
    if(!cat) return;
    const pos=screenToFlowPosition({ x:e.clientX,y:e.clientY });
    if(cat==="area"){
      setNodes(ns=>[...ns,{ id:uid("area"),type:"area",position:pos,style:{ width:260,height:180 },
        data:{ areaType:sub,label:"",summary:"",requirements:[],autoInlets:[],autoOutlets:[],handles:["top","bottom","left","right"] },zIndex:-1 }]);
    } else if(cat==="equipment"){
      const def=EQUIP_DEFAULTS[sub]||{};
      setNodes(ns=>[...ns,{ id:uid("eq"),type:"equipment",position:pos,
        data:{ equipType:sub,itemNo:"",label:sub,handles:["top","bottom","left","right"],requirements:[],...def } }]);
    } else if(cat==="instrument"){
      const [ic,it]=sub.split("|");
      setNodes(ns=>[...ns,{ id:uid("ins"),type:"instrument",position:pos,
        data:{ instrCategory:ic,instrType:it,itemNo:"",requirements:[] } }]);
    } else if(cat==="connection"){
      // Brench/Process Gas/Material 드롭 → 분기점 노드 생성
      if(sub==="Piping"||sub==="Duct"||sub==="Conveyor"){
        // Piping/Duct는 핸들 드래그로 연결 — 드롭 시 Brench 생성
        setNodes(ns=>[...ns,{ id:uid("br"),type:"brench",position:pos,data:{ _hint:sub } }]);
      } else {
        setNodes(ns=>[...ns,{ id:uid("br"),type:"brench",position:pos,data:{ _hint:sub } }]);
      }
    }
  },[screenToFlowPosition,setNodes]);

  // CONNECT — 드래그 힌트 lineType 감지
  const onConnect=useCallback(params=>{
    connRef.current=params;
    // source 노드의 _hint 로 default lineType 결정
    const srcNode=params.source;
    // 기본값 Piping
    setModalDefault("Piping");
    setModal(true);
  },[]);

  const confirmConn=useCallback(lineType=>{
    const p=connRef.current; if(!p) return;
    setEdges(es=>addEdge({ ...p,id:uid("e"),type:"pipe",data:{ lineType } },es));
    setModal(false); connRef.current=null;
  },[setEdges]);

  const onNodeClick=useCallback((_,n)=>setSel(n),[]);
  const onEdgeClick=useCallback((_,e)=>setSel(e),[]);
  const onPaneClick=useCallback(()=>setSel(null),[]);

  const onUpdateNode=useCallback((id,newData)=>{
    setNodes(ns=>ns.map(n=>n.id===id?{...n,data:newData}:n));
    setSel(prev=>prev&&prev.id===id?{...prev,data:newData}:prev);
  },[setNodes]);
  const onUpdateEdge=useCallback((id,newData)=>{
    setEdges(es=>es.map(e=>e.id===id?{...e,data:newData}:e));
    setSel(prev=>prev&&prev.id===id?{...prev,data:newData}:prev);
  },[setEdges]);
  const onAddHandle=useCallback((nodeId,dir)=>{
    setNodes(ns=>ns.map(n=>{ if(n.id!==nodeId) return n; return { ...n,data:{ ...n.data,handles:[...(n.data.handles||["top","bottom","left","right"]),dir] } }; }));
  },[setNodes]);
  const onDeleteSel=useCallback(()=>{
    if(!sel) return;
    if(sel.position){ setNodes(ns=>ns.filter(n=>n.id!==sel.id)); setEdges(es=>es.filter(e=>e.source!==sel.id&&e.target!==sel.id)); }
    else { setEdges(es=>es.filter(e=>e.id!==sel.id)); }
    setSel(null);
  },[sel,setNodes,setEdges]);

  // 키보드 단축키
  useEffect(()=>{
    const fn=e=>{
      const tag=document.activeElement?.tagName?.toLowerCase();
      const inInput=tag==="input"||tag==="textarea"||tag==="select";
      if((e.key==="Delete"||e.key==="Backspace")&&sel&&!inInput){ onDeleteSel(); return; }
      // Undo
      if(e.ctrlKey&&e.key==="z"&&!inInput){ e.preventDefault(); undo(); return; }
      // Redo
      if((e.ctrlKey&&e.key==="y"||(e.ctrlKey&&e.shiftKey&&e.key==="Z"))&&!inInput){ e.preventDefault(); redo(); return; }
      if(e.ctrlKey&&e.key==="c"&&sel?.position){ clipRef.current=sel; return; }
      if(e.ctrlKey&&e.key==="v"&&clipRef.current){
        const src=clipRef.current;
        const newNode={
          ...src, id:uid(src.type),
          position:{ x:src.position.x+30,y:src.position.y+30 },
          data:{ ...src.data,itemNo:src.data.itemNo?(src.data.itemNo+" copy"):"",autoInlets:[],autoOutlets:[],requirements:[] },
          selected:false,
        };
        setNodes(ns=>[...ns,newNode]); setSel(newNode); clipRef.current=newNode; return;
      }
    };
    window.addEventListener("keydown",fn);
    return()=>window.removeEventListener("keydown",fn);
  },[sel,onDeleteSel,setNodes,undo,redo]);

  // sel 동기화
  useEffect(()=>{
    if(!sel?.position) return;
    const fresh=nodes.find(n=>n.id===sel.id);
    if(fresh&&fresh!==sel) setSel(fresh);
  },[nodes]); // eslint-disable-line

  // JSON Export / Import
  const onExport=()=>{
    const blob=new Blob([JSON.stringify({nodes,edges},null,2)],{type:"application/json"});
    const a=document.createElement("a"); a.href=URL.createObjectURL(blob);
    a.download=`MBSE_${new Date().toISOString().slice(0,10)}.json`; a.click();
  };
  const onImport=e=>{
    const f=e.target.files[0]; if(!f) return;
    const r=new FileReader();
    r.onload=ev=>{ try{ const p=JSON.parse(ev.target.result); if(p.nodes)setNodes(p.nodes); if(p.edges)setEdges(p.edges); }catch{ alert("Invalid JSON"); } };
    r.readAsText(f); e.target.value="";
  };

  // Excel Export
  const onExcelExport = async () => {
    try {
      setXlsxMsg("Excel 생성 중...");
      await exportToExcel(nodes, edges);
      setXlsxMsg("");
    } catch(err) {
      setXlsxMsg("오류: " + err.message);
      setTimeout(()=>setXlsxMsg(""), 3000);
    }
  };

  // Excel Import
  const onExcelImport = async (e) => {
    const f = e.target.files[0]; if(!f) return;
    try {
      setXlsxMsg("Excel 불러오는 중...");
      await importFromExcel(f, nodes, edges, setNodes, setEdges);
      setXlsxMsg("✅ Import 완료!");
      setTimeout(()=>setXlsxMsg(""), 2500);
    } catch(err) {
      setXlsxMsg("오류: " + err.message);
      setTimeout(()=>setXlsxMsg(""), 4000);
    }
    e.target.value="";
  };

  return (
    <>
    <style>{GLOBAL_CSS}</style>
    <div style={{ display:"flex",height:"100vh",width:"100vw",fontFamily:"'Segoe UI',sans-serif",background:"#f8fafc",overflow:"hidden" }}>
      <Sidebar onDragStart={onDragStart}/>
      <div style={{ flex:1,display:"flex",flexDirection:"column",minWidth:0 }}>
        <div style={{ height:44,background:"#0f172a",display:"flex",alignItems:"center",padding:"0 16px",gap:8,flexShrink:0 }}>
          <span style={{ color:"#f1f5f9",fontWeight:800,fontSize:14,marginRight:6,letterSpacing:0.5 }}>⬡ MBSE Interface Master</span>

          {/* ↩ Undo / ↪ Redo */}
          <button onClick={undo}
            title="되돌리기 (Ctrl+Z)"
            style={{ background:"#1e293b",color: undoStack.current.length>0?"#94a3b8":"#334155",border:"1px solid #334155",borderRadius:5,padding:"3px 10px",cursor: undoStack.current.length>0?"pointer":"default",fontSize:12 }}>
            ↩
          </button>
          <button onClick={redo}
            title="다시하기 (Ctrl+Y)"
            style={{ background:"#1e293b",color: redoStack.current.length>0?"#94a3b8":"#334155",border:"1px solid #334155",borderRadius:5,padding:"3px 10px",cursor: redoStack.current.length>0?"pointer":"default",fontSize:12 }}>
            ↪
          </button>

          {/* 구분선 */}
          <div style={{ width:1,height:20,background:"#334155",margin:"0 2px" }}/>

          {/* 💾 Save 버튼 */}
          <button onClick={onSave}
            style={{ background:"#1d4ed8",color:"#fff",border:"none",borderRadius:5,padding:"3px 12px",cursor:"pointer",fontSize:11,fontWeight:700 }}>
            💾 Save
          </button>

          {/* 구분선 */}
          <div style={{ width:1,height:20,background:"#334155",margin:"0 2px" }}/>

          {/* JSON */}
          <button onClick={onExport} style={{ background:"#1e293b",color:"#94a3b8",border:"1px solid #334155",borderRadius:5,padding:"3px 10px",cursor:"pointer",fontSize:11 }}>⬇ JSON</button>
          <button onClick={()=>fileRef.current?.click()} style={{ background:"#1e293b",color:"#94a3b8",border:"1px solid #334155",borderRadius:5,padding:"3px 10px",cursor:"pointer",fontSize:11 }}>⬆ JSON</button>
          <input ref={fileRef} type="file" accept=".json" style={{ display:"none" }} onChange={onImport}/>

          {/* 구분선 */}
          <div style={{ width:1,height:20,background:"#334155",margin:"0 2px" }}/>

          {/* 전체보기 */}
          <button onClick={()=>fitView({ padding:0.15, duration:400 })}
            style={{ background:"#1e293b",color:"#94a3b8",border:"1px solid #334155",borderRadius:5,padding:"3px 10px",cursor:"pointer",fontSize:11 }}>
            ⊞ 전체보기
          </button>

          {/* Excel */}
          <button onClick={onExcelExport} style={{ background:"#14532d",color:"#86efac",border:"1px solid #166534",borderRadius:5,padding:"3px 10px",cursor:"pointer",fontSize:11,fontWeight:600 }}>📊 Excel ⬇</button>
          <button onClick={()=>xlsxRef.current?.click()} style={{ background:"#14532d",color:"#86efac",border:"1px solid #166534",borderRadius:5,padding:"3px 10px",cursor:"pointer",fontSize:11,fontWeight:600 }}>📊 Excel ⬆</button>
          <input ref={xlsxRef} type="file" accept=".xlsx,.xls" style={{ display:"none" }} onChange={onExcelImport}/>

          {/* 구분선 */}
          <div style={{ width:1,height:20,background:"#334155",margin:"0 2px" }}/>

          {/* 📋 History Log 버튼 */}
          <button onClick={()=>setShowLog(v=>!v)}
            style={{ background: showLog?"#0f172a":"#1e293b", color: showLog?"#fbbf24":"#94a3b8", border:`1px solid ${showLog?"#fbbf24":"#334155"}`, borderRadius:5, padding:"3px 10px", cursor:"pointer", fontSize:11 }}>
            📋 Log {history.length>0 && <span style={{ background:"#fbbf24",color:"#0f172a",borderRadius:8,padding:"0 5px",fontSize:10,fontWeight:700,marginLeft:3 }}>{history.length}</span>}
          </button>

          {/* 상태 메시지 */}
          {(xlsxMsg||saveMsg) && (
            <span style={{ color:"#86efac",fontSize:11,marginLeft:4 }}>{xlsxMsg||saveMsg}</span>
          )}

          <span style={{ marginLeft:"auto",color:"#475569",fontSize:11 }}>{nodes.length} nodes · {edges.length} edges</span>
        </div>
        <div style={{ flex:1,display:"flex",minHeight:0 }}>
          <div ref={wrapRef} style={{ flex:1,minWidth:0 }} onDragOver={onDragOver} onDrop={onDrop}>
            <ReactFlow
              nodes={nodes} edges={edges}
              onNodesChange={onNodesChange} onEdgesChange={onEdgesChange}
              onConnect={onConnect}
              onNodeClick={onNodeClick} onEdgeClick={onEdgeClick} onPaneClick={onPaneClick}
              // ── 엣지 경로 수동 조정 핸들러 ──────────────────
              onEdgeUpdate={(oldEdge, newConnection) => {
                setEdges(es => es.map(e => e.id === oldEdge.id ? { ...e, ...newConnection } : e));
              }}
              onEdgeUpdateStart={() => {}}
              onEdgeUpdateEnd={() => {}}
              nodeTypes={nodeTypes} edgeTypes={edgeTypes}
              connectionMode={ConnectionMode.Loose}
              connectionLineType="smoothstep"
              connectionLineStyle={{ stroke:"#3b82f6",strokeWidth:2 }}
              defaultEdgeOptions={{ type:"pipe" }}
              // ── fitView: 전체 노드가 화면에 맞게 보임 ─────────
              fitView
              fitViewOptions={{ padding:0.15, includeHiddenNodes:false }}
              minZoom={0.05}
              maxZoom={2}
              snapToGrid snapGrid={[10,10]}
              deleteKeyCode={null}
              // ── 엣지 경로 수동 조정 가능 ─────────────────────
              edgesUpdatable={true}
              edgeUpdaterRadius={12}
              // ── 캔버스 영역 8배 확장 (±16000px) ─────────────
              translateExtent={[[-16000,-16000],[16000,16000]]}
              defaultViewport={{ x:0, y:0, zoom:0.5 }}
              // ── 범위 선택 (드래그로 여러 노드 선택) ──────────
              selectionMode="partial"
              selectionOnDrag={true}
              panOnDrag={[1,2]}
              multiSelectionKeyCode="Shift"
              selectionKeyCode={null}
            >
              <Controls/>
              <MiniMap nodeColor={n=>n.type==="instrument"?"#a855f7":n.type==="area"?"#93c5fd":"#3b82f6"} maskColor="rgba(0,0,0,0.04)"/>
              <Background variant="dots" gap={20} size={1} color="#cbd5e1"/>
              <Panel position="bottom-left">
                <div style={{ background:"rgba(255,255,255,0.92)",border:"1px solid #e2e8f0",borderRadius:7,padding:"5px 10px",fontSize:10,color:"#64748b" }}>
                  드래그 → 범위선택 · Shift+클릭 → 추가선택 · Ctrl+Z 되돌리기 · Ctrl+C/V 복사붙여넣기 · Del 삭제
                </div>
              </Panel>
            </ReactFlow>
          </div>
          <Inspector sel={sel} nodes={nodes} edges={edges} onUpdateNode={onUpdateNode} onUpdateEdge={onUpdateEdge} onDeleteSel={onDeleteSel} onAddHandle={onAddHandle}/>

          {/* ── History Log 패널 ── */}
          {showLog && (
            <div style={{ width:280,background:"#0f172a",borderLeft:"1px solid #1e293b",display:"flex",flexDirection:"column",flexShrink:0 }}>
              {/* 헤더 */}
              <div style={{ padding:"10px 14px 8px",borderBottom:"1px solid #1e293b",display:"flex",justifyContent:"space-between",alignItems:"center",flexShrink:0 }}>
                <span style={{ fontWeight:800,fontSize:13,color:"#f1f5f9" }}>📋 저장 이력</span>
                <div style={{ display:"flex",gap:6 }}>
                  <button onClick={()=>{
                    if(window.confirm("저장 이력을 전체 삭제할까요?")) {
                      setHistory([]);
                      localStorage.removeItem("mbse_autosave");
                    }
                  }} style={{ background:"#1e293b",color:"#94a3b8",border:"1px solid #334155",borderRadius:4,padding:"2px 8px",cursor:"pointer",fontSize:10 }}>
                    전체삭제
                  </button>
                  <button onClick={()=>setShowLog(false)}
                    style={{ background:"#1e293b",color:"#94a3b8",border:"1px solid #334155",borderRadius:4,padding:"2px 8px",cursor:"pointer",fontSize:10 }}>✕</button>
                </div>
              </div>

              {/* 안내 */}
              <div style={{ padding:"6px 14px",borderBottom:"1px solid #1e293b",fontSize:10,color:"#475569",flexShrink:0 }}>
                💾 Save 버튼을 누를 때만 이력이 기록됩니다.
              </div>

              {/* 로그 목록 */}
              <div style={{ overflowY:"auto",flex:1,padding:"6px 0" }}>
                {history.length===0 && (
                  <div style={{ padding:"20px 14px",color:"#475569",fontSize:11,textAlign:"center",lineHeight:1.8 }}>
                    저장 이력이 없습니다.<br/>
                    <span style={{ color:"#1d4ed8" }}>💾 Save</span> 버튼을 누르면<br/>이력이 기록됩니다.
                  </div>
                )}
                {history.map((log,i)=>(
                  <div key={log.id} style={{
                    padding:"10px 14px",
                    borderBottom:"1px solid #1e293b",
                    background: i===0?"rgba(29,78,216,0.12)":"transparent",
                  }}>
                    {/* 최신 뱃지 */}
                    {i===0 && (
                      <div style={{ display:"inline-block",fontSize:9,background:"#1d4ed8",color:"#fff",borderRadius:3,padding:"0 6px",marginBottom:5,fontWeight:700 }}>
                        최신
                      </div>
                    )}
                    {/* 3가지 핵심 정보 */}
                    <div style={{ display:"flex",flexDirection:"column",gap:3,marginBottom:7 }}>
                      <div style={{ display:"flex",alignItems:"center",gap:5 }}>
                        <span style={{ fontSize:9,color:"#475569",width:14 }}>📅</span>
                        <span style={{ fontSize:11,color:"#94a3b8" }}>{log.date}</span>
                      </div>
                      <div style={{ display:"flex",alignItems:"center",gap:5 }}>
                        <span style={{ fontSize:9,color:"#475569",width:14 }}>👤</span>
                        <span style={{ fontSize:12,color:"#e2e8f0",fontWeight:600 }}>{log.author}</span>
                      </div>
                      <div style={{ display:"flex",alignItems:"flex-start",gap:5 }}>
                        <span style={{ fontSize:9,color:"#475569",width:14,marginTop:1 }}>📝</span>
                        <span style={{ fontSize:11,color:"#cbd5e1",lineHeight:1.5,flex:1 }}>
                          {log.summary.length>40 ? log.summary.slice(0,40)+"…" : log.summary}
                        </span>
                      </div>
                    </div>
                    {/* 상세보기 버튼 */}
                    <button
                      onClick={()=>setLogDetail(log)}
                      style={{ width:"100%",background:"#1e293b",color:"#94a3b8",border:"1px solid #334155",borderRadius:5,padding:"4px 0",cursor:"pointer",fontSize:10,fontWeight:600 }}>
                      🔍 상세보기
                    </button>
                  </div>
                ))}
              </div>

              {/* 하단 */}
              <div style={{ padding:"8px 14px",borderTop:"1px solid #1e293b",fontSize:10,color:"#475569",flexShrink:0 }}>
                총 {history.length}건 · 노드 {nodes.length}개 · 연결 {edges.length}개
              </div>
            </div>
          )}
        </div>
      </div>
      {modal&&<ConnModal defaultType={modalDefault} onConfirm={confirmConn} onCancel={()=>{ setModal(false); connRef.current=null; }}/>}

      {/* 저장 메모 모달 */}
      {showSaveModal && (
        <SaveMemoModal
          nodes={nodes} edges={edges}
          onConfirm={confirmSave}
          onCancel={()=>setShowSaveModal(false)}
        />
      )}

      {/* 로그 상세보기 모달 */}
      {logDetail && <LogDetailModal log={logDetail} onClose={()=>setLogDetail(null)}/>}
    </div>
    </>
  );
};

export default function App() {
  return <ReactFlowProvider><CanvasInner/></ReactFlowProvider>;
}
