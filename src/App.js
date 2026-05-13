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
} from "reactflow";
import "reactflow/dist/style.css";

// ─────────────────────────────────────────────────────────────
// xlsx-js-style LOADER  (스타일 지원 SheetJS fork)
// CDN: unpkg.com/xlsx-js-style
// ─────────────────────────────────────────────────────────────
const loadXLSX = () => new Promise((resolve, reject) => {
  if (window.XLSXStyle) { resolve(window.XLSXStyle); return; }
  const script = document.createElement("script");
  script.src = "https://unpkg.com/xlsx-js-style@1.2.0/dist/xlsx.bundle.js";
  script.onload  = () => {
    if (window.XLSXStyle) { resolve(window.XLSXStyle); return; }
    // fallback: 일부 빌드는 window.XLSX로 노출
    if (window.XLSX) { window.XLSXStyle = window.XLSX; resolve(window.XLSX); return; }
    reject(new Error("xlsx-js-style 로드 실패"));
  };
  script.onerror = () => {
    // fallback to plain SheetJS
    if (window.XLSX) { resolve(window.XLSX); return; }
    const s2 = document.createElement("script");
    s2.src = "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";
    s2.onload  = () => resolve(window.XLSX);
    s2.onerror = () => reject(new Error("SheetJS 로드 실패"));
    document.head.appendChild(s2);
  };
  document.head.appendChild(script);
});

// ─────────────────────────────────────────────────────────────
// 스타일 헬퍼 (xlsx-js-style 전용)
// ─────────────────────────────────────────────────────────────
const XS = {
  // 셀 스타일 생성
  s: (opts = {}) => ({
    font:      { name:"Arial", sz: opts.sz||9, bold:!!opts.bold,
                 color:{ rgb: opts.fc || "000000" } },
    fill:      opts.bg
                 ? { patternType:"solid", fgColor:{ rgb: opts.bg } }
                 : { patternType:"none" },
    alignment: { horizontal: opts.h||"center", vertical:"center",
                 wrapText: !!opts.wrap },
    border:    opts.border===false ? {} : {
      top:    { style:"thin", color:{ rgb: opts.bc||"BFBFBF" } },
      bottom: { style:"thin", color:{ rgb: opts.bc||"BFBFBF" } },
      left:   { style:"thin", color:{ rgb: opts.bc||"BFBFBF" } },
      right:  { style:"thin", color:{ rgb: opts.bc||"BFBFBF" } },
    },
  }),
  // 셀 객체
  c: (v, opts={}) => ({ v, s: XS.s(opts) }),
  // 헤더 셀 (네이비 배경 흰 글씨)
  hdr: (v) => XS.c(v,  { bold:true, bg:"1F3864", fc:"FFFFFF", bc:"1F3864" }),
  // 서브 헤더 (파란 배경 흰 글씨)
  shdr:(v) => XS.c(v,  { bold:true, bg:"2E75B6", fc:"FFFFFF", bc:"2E75B6" }),
  // 상태 셀
  status: (v) => {
    const map = {
      "OPEN":        { bg:"FFF2CC", fc:"7F6000" },
      "IN PROGRESS": { bg:"DDEBF7", fc:"1F4E79" },
      "CLOSED":      { bg:"E2EFDA", fc:"375623" },
      "OVERDUE":     { bg:"FCE4D6", fc:"843C0C" },
    };
    const m = map[v] || { bg:"F2F2F2", fc:"595959" };
    return XS.c(v, { bold:true, ...m });
  },
  // 우선순위 셀
  priority: (v) => {
    const map = {
      "Critical":{ fc:"DC2626" },
      "High":    { fc:"EA580C" },
      "Medium":  { fc:"CA8A04" },
      "Low":     { fc:"16A34A" },
    };
    const m = map[v] || {};
    return XS.c(v, { bold:true, ...m });
  },
  // 번갈아 배경 (짝수 행 연파랑)
  alt: (v, i, opts={}) => XS.c(v, { bg: i%2===0 ? "FFFFFF":"F2F7FC", ...opts }),
  // 컬럼 너비 자동 계산
  colWidths: (headers, rows) =>
    headers.map(h => ({
      wch: Math.min(40, Math.max(
        h.length+2,
        ...rows.map(r => String(r[h]||"").length),
        8
      ))
    })),
};

// ─────────────────────────────────────────────────────────────
// EXCEL EXPORT — IC Register + ICD + Equipment + Requirements
// ─────────────────────────────────────────────────────────────
const exportToExcel = async (nodes, edges) => {
  const XLSX = await loadXLSX();

  // ── 유틸: 셀 스타일 헬퍼 ──────────────────────────────────
  // SheetJS CE는 스타일을 직접 지원하지 않으므로
  // xlsx-js-style CDN을 동적 로드해서 스타일 적용
  // (없으면 기본 포맷으로 출력)

  // ── Area 포함 노드 계산 ───────────────────────────────────
  const getAreaOf = (node, allNodes) => {
    return allNodes.find(a => {
      if (a.type !== "area") return false;
      const ax = a.position.x, ay = a.position.y;
      const aw = a.style?.width || a.width || 260;
      const ah = a.style?.height || a.height || 180;
      return (
        node.position.x >= ax && node.position.y >= ay &&
        node.position.x <= ax + aw && node.position.y <= ay + ah
      );
    });
  };

  // ── IC 번호 자동 생성 유틸 ────────────────────────────────
  let icCounter = 1;
  const nextIC = () => `IC-${String(icCounter++).padStart(3,"0")}`;

  // ══════════════════════════════════════════════════════════
  // SHEET 1: IC Register
  // 경계를 넘는 모든 Connection → IC 항목으로 변환
  // ══════════════════════════════════════════════════════════
  const areaNodes = nodes.filter(n => n.type === "area");
  const icRows = [];

  // 각 Area별 경계 통과 엣지 수집
  areaNodes.forEach(area => {
    const ax = area.position.x, ay = area.position.y;
    const aw = area.style?.width || area.width || 260;
    const ah = area.style?.height || area.height || 180;
    const areaLabel = area.data?.label
      ? `[${area.data.areaType}] ${area.data.label}`
      : `[${area.data.areaType}]`;

    const insideIds = new Set(
      nodes.filter(n => {
        if (n.type === "area") return false;
        return n.position.x >= ax && n.position.y >= ay &&
               n.position.x <= ax+aw && n.position.y <= ay+ah;
      }).map(n => n.id)
    );
    if (insideIds.size === 0) return;

    edges.forEach(e => {
      const srcIn = insideIds.has(e.source);
      const tgtIn = insideIds.has(e.target);
      if (srcIn === tgtIn) return; // 경계 통과 아님

      const d = e.data || {};
      const srcNode = nodes.find(n => n.id === e.source);
      const tgtNode = nodes.find(n => n.id === e.target);
      const srcLabel = srcNode?.data?.itemNo || srcNode?.data?.label || e.source;
      const tgtLabel = tgtNode?.data?.itemNo || tgtNode?.data?.label || e.target;

      const srcArea = getAreaOf(srcNode || {position:{x:-9999,y:-9999}}, nodes);
      const tgtArea = getAreaOf(tgtNode || {position:{x:-9999,y:-9999}}, nodes);

      const fromOrg = srcArea
        ? (srcArea.data?.label
            ? `[${srcArea.data.areaType}] ${srcArea.data.label}`
            : `[${srcArea.data.areaType}]`)
        : "외부";
      const toOrg = tgtArea
        ? (tgtArea.data?.label
            ? `[${tgtArea.data.areaType}] ${tgtArea.data.label}`
            : `[${tgtArea.data.areaType}]`)
        : "외부";

      const fluidSub  = d.fluidSub    || "";
      const sizeVal   = d.sizeNum ? `${d.sizeNum}A` : (d.size || "");
      const lineLabel = fluidSub && sizeVal
        ? `${fluidSub}-${sizeVal}`
        : fluidSub || sizeVal || d.lineType || "Piping";

      const dir = tgtIn ? "INLET" : "OUTLET";

      icRows.push({
        "IC No.":           nextIC(),
        "인터페이스 제목":   `${lineLabel} — ${fromOrg} → ${toOrg}`,
        "분류":             d.lineType === "Process Gas" ? "Process"
                           : d.lineType === "Duct"       ? "Mechanical"
                           : d.lineType === "Conveyor"   ? "Mechanical"
                           : "Process",
        "발신 조직 (From)": fromOrg,
        "수신 조직 (To)":   toOrg,
        "관련 Package":     areaLabel,
        "인터페이스 설명":  `${dir}: ${lineLabel} (${srcLabel} → ${tgtLabel})`,
        "유체/매체":        d.fluidPrimary
                             ? `${d.fluidPrimary} / ${fluidSub}`
                             : (d.lineType || ""),
        "Size":             sizeVal,
        "Schedule":         d.spec      || "",
        "Line No.":         d.serialNo  || "",
        "Line Text":        d.lineText  || "",
        "From 설비":        srcLabel,
        "To 설비":          tgtLabel,
        "상태":             "OPEN",
        "우선순위":         fluidSub === "" ? "Medium" : "High",
        "등록일":           new Date().toISOString().slice(0,10),
        "목표 완료일":      "",
        "실제 완료일":      "",
        "담당자 (From)":    "",
        "담당자 (To)":      "",
        "비고":             "",
        "ICD 번호":         `ICD-${String(icRows.length+1).padStart(3,"0")}`,
        "Edge ID":          e.id,
      });
    });
  });

  // Area 없는 경계 통과 엣지도 추가
  const allAreaInsideIds = new Set(
    nodes.filter(n => {
      if (n.type === "area") return false;
      return areaNodes.some(a => {
        const ax=a.position.x, ay=a.position.y;
        const aw=a.style?.width||a.width||260;
        const ah=a.style?.height||a.height||180;
        return n.position.x>=ax&&n.position.y>=ay&&
               n.position.x<=ax+aw&&n.position.y<=ay+ah;
      });
    }).map(n => n.id)
  );
  edges.forEach(e => {
    if (icRows.find(r => r["Edge ID"] === e.id)) return;
    const d = e.data || {};
    const srcNode = nodes.find(n => n.id === e.source);
    const tgtNode = nodes.find(n => n.id === e.target);
    const srcLabel = srcNode?.data?.itemNo||srcNode?.data?.label||e.source;
    const tgtLabel = tgtNode?.data?.itemNo||tgtNode?.data?.label||e.target;
    const fluidSub = d.fluidSub || "";
    const sizeVal  = d.sizeNum ? `${d.sizeNum}A` : (d.size || "");
    const lineLabel = fluidSub && sizeVal
      ? `${fluidSub}-${sizeVal}` : fluidSub||sizeVal||d.lineType||"Piping";
    icRows.push({
      "IC No.":           nextIC(),
      "인터페이스 제목":   `${lineLabel} — ${srcLabel} → ${tgtLabel}`,
      "분류":             "Process",
      "발신 조직 (From)": srcLabel,
      "수신 조직 (To)":   tgtLabel,
      "관련 Package":     "",
      "인터페이스 설명":  `${lineLabel} (${srcLabel} → ${tgtLabel})`,
      "유체/매체":        d.fluidPrimary ? `${d.fluidPrimary} / ${fluidSub}` : (d.lineType||""),
      "Size":             sizeVal,
      "Schedule":         d.spec||"",
      "Line No.":         d.serialNo||"",
      "Line Text":        d.lineText||"",
      "From 설비":        srcLabel,
      "To 설비":          tgtLabel,
      "상태":             "OPEN",
      "우선순위":         "Medium",
      "등록일":           new Date().toISOString().slice(0,10),
      "목표 완료일":      "",
      "실제 완료일":      "",
      "담당자 (From)":    "",
      "담당자 (To)":      "",
      "비고":             "",
      "ICD 번호":         `ICD-${String(icRows.length+1).padStart(3,"0")}`,
      "Edge ID":          e.id,
    });
  });

  // ══════════════════════════════════════════════════════════
  // SHEET 2: Equipment List (설비 목록)
  // ══════════════════════════════════════════════════════════
  const equipRows = nodes
    .filter(n => n.type === "equipment" || n.type === "instrument")
    .map(n => {
      const d = n.data || {};
      const area = getAreaOf(n, nodes);
      const areaLabel = area
        ? (area.data?.label
            ? `[${area.data.areaType}] ${area.data.label}`
            : `[${area.data.areaType}]`)
        : "";
      // 연결된 엣지 수
      const connEdges = edges.filter(e => e.source===n.id||e.target===n.id);
      return {
        "Item No.":     d.itemNo      || "",
        "설비명":        d.label       || d.equipType || d.instrCategory || "",
        "설비 유형":     n.type === "instrument"
                          ? `Instrument (${d.instrCategory||""} ${d.instrType||""})`
                          : (d.equipType || ""),
        "소속 Area":     areaLabel,
        "재질":          d.material    || "",
        "용량":          d.capacity    || "",
        "설계 압력":     d.designP     || "",
        "설계 온도":     d.designT     || "",
        "연결 Interface 수": connEdges.length,
        "비고":          d.summary     || "",
        "Node ID":       n.id,
      };
    });

  // ══════════════════════════════════════════════════════════
  // SHEET 3: Connection List (배관 라인 목록)
  // ══════════════════════════════════════════════════════════
  const connRows = edges.map(e => {
    const d = e.data || {};
    const srcNode = nodes.find(n => n.id === e.source);
    const tgtNode = nodes.find(n => n.id === e.target);
    const srcLabel = srcNode?.data?.itemNo||srcNode?.data?.label||e.source;
    const tgtLabel = tgtNode?.data?.itemNo||tgtNode?.data?.label||e.target;
    const sizeVal  = d.sizeNum ? `${d.sizeNum}A` : (d.size||"");
    const icMatch  = icRows.find(r => r["Edge ID"] === e.id);
    return {
      "Line No.":         d.serialNo   || "",
      "Line Type":        d.lineType   || "Piping",
      "Fluid (Primary)":  d.fluidPrimary|| "",
      "Fluid (Sub)":      d.fluidSub   || "",
      "Size":             sizeVal,
      "Schedule":         d.spec       || "",
      "Line Text":        d.lineText   || "",
      "From (Item No.)":  srcLabel,
      "To (Item No.)":    tgtLabel,
      "연결 IC No.":       icMatch?.["IC No."] || "",
      "IC 상태":           icMatch?.["상태"]   || "",
      "Edge ID":           e.id,
    };
  });

  // ══════════════════════════════════════════════════════════
  // SHEET 4: Requirements (요구사항)
  // ══════════════════════════════════════════════════════════
  const reqRows = [];
  nodes.forEach(n => {
    const d = n.data || {};
    const area = getAreaOf(n, nodes);
    const areaLabel = area
      ? (area.data?.label
          ? `[${area.data.areaType}] ${area.data.label}`
          : `[${area.data.areaType}]`)
      : "";
    (d.requirements || []).forEach(r => {
      reqRows.push({
        "Node ID":      n.id,
        "Item No.":     d.itemNo || d.label || "",
        "설비 유형":     n.type,
        "소속 Area":     areaLabel,
        "Stakeholder":  r.who  || "",
        "날짜":          r.date || "",
        "요구사항":       r.text || "",
      });
    });
  });

  // ══════════════════════════════════════════════════════════
  // 워크북 조립 — 서식 적용
  // ══════════════════════════════════════════════════════════
  const wb = XLSX.utils.book_new();
  const today = new Date().toISOString().slice(0,10);
  const toDate = new Date().toLocaleString("ko-KR",{
    year:"numeric",month:"2-digit",day:"2-digit",
    hour:"2-digit",minute:"2-digit"
  });

  // ── 공통: aoa → ws 변환 헬퍼 ──────────────────────────────
  const aoaToSheet = (aoa) => {
    const ws = {};
    let maxC = 0;
    aoa.forEach((row, R) => {
      maxC = Math.max(maxC, row.length);
      row.forEach((cell, C) => {
        if (cell == null) return;
        const ref = XLSX.utils.encode_cell({r:R, c:C});
        if (typeof cell === "object" && "v" in cell) {
          ws[ref] = cell;
        } else {
          ws[ref] = { v: cell, s: XS.s({ h:"left" }) };
        }
      });
    });
    ws["!ref"] = XLSX.utils.encode_range({
      s:{r:0,c:0}, e:{r:aoa.length-1, c:maxC-1}
    });
    return ws;
  };

  // ── SHEET 1: IC Register ───────────────────────────────────
  const icHdrs = [
    "IC No.","인터페이스 제목","분류","발신 조직 (From)","수신 조직 (To)",
    "관련 Package","인터페이스 설명","유체/매체","Size","Schedule",
    "Line No.","From 설비","To 설비","상태","우선순위",
    "등록일","목표 완료일","실제 완료일","담당자 (From)","담당자 (To)",
    "비고","ICD 번호",
  ];

  // 프로젝트 정보 행
  const icAoa = [
    // Row 0: 타이틀
    [Object.assign(XS.hdr("IC&TI — Interface Check Register"), {
      s: XS.s({ bold:true, sz:13, bg:"1F3864", fc:"FFFFFF", bc:"1F3864" })
    }), ...Array(icHdrs.length-1).fill(null)],
    // Row 1: 프로젝트 정보
    [XS.c("Project", {bold:true,bg:"D9E1F2",fc:"1F3864",h:"left"}),
     XS.c("수소환원제철 Plant — HyREX Project",{bg:"FFFFFF",h:"left"}),
     null,null,
     XS.c("작성일", {bold:true,bg:"D9E1F2",fc:"1F3864",h:"left"}),
     XS.c(toDate,   {bg:"FFFFFF",h:"left"}),
     null,null,
     XS.c("총 IC 건수", {bold:true,bg:"D9E1F2",fc:"1F3864",h:"left"}),
     XS.c(`${icRows.length}건`, {bold:true,bg:"FFFFFF",fc:"1D4ED8"}),
     null,null,null,null,null,null,null,null,null,null,null,null],
    // Row 2: 범례
    [Object.assign(
      XS.c("상태 범례:   OPEN = 미해결     IN PROGRESS = 진행중     CLOSED = 완료     OVERDUE = 지연",
        { h:"left", bg:"F8F8F8", fc:"595959" }),
      { s: { ...XS.s({h:"left",bg:"F8F8F8",fc:"595959"}),
             font:{ name:"Arial",sz:8,italic:true,color:{rgb:"595959"} } } }
    ), ...Array(icHdrs.length-1).fill(null)],
    // Row 3: 헤더
    icHdrs.map(h => XS.shdr(h)),
    // 데이터 행
    ...icRows.map((row, i) => icHdrs.map((h, j) => {
      const v = row[h] ?? "";
      if (h === "상태")    return XS.status(v);
      if (h === "우선순위") return XS.priority(v);
      if (h === "IC No."||h==="ICD 번호"||h==="Line No.")
        return XS.alt(v, i, { bold:true, fc:"1D4ED8" });
      if (h === "인터페이스 제목"||h==="인터페이스 설명")
        return XS.alt(v, i, { h:"left", wrap:true });
      return XS.alt(v, i, { h: j<2?"center":"left" });
    })),
    // 요약 행
    [XS.c("집계", {bold:true,bg:"1F3864",fc:"FFFFFF"}),
     XS.c(`전체 ${icRows.length}건`, {bold:true,bg:"2E75B6",fc:"FFFFFF"}),
     null,null,null,null,null,null,null,null,null,null,null,
     XS.c(`OPEN: ${icRows.filter(r=>r["상태"]==="OPEN").length}`,
       {bold:true,bg:"FFF2CC",fc:"7F6000"}),
     XS.c(`HIGH: ${icRows.filter(r=>r["우선순위"]==="High"||r["우선순위"]==="Critical").length}`,
       {bold:true,bg:"FCE4D6",fc:"843C0C"}),
     ...Array(icHdrs.length-15).fill(null)],
  ];

  const wsIC = aoaToSheet(icAoa);
  // 열 병합 (타이틀, 범례)
  wsIC["!merges"] = [
    { s:{r:0,c:0}, e:{r:0,c:icHdrs.length-1} },
    { s:{r:2,c:0}, e:{r:2,c:icHdrs.length-1} },
    { s:{r:1,c:1}, e:{r:1,c:3} },
    { s:{r:1,c:5}, e:{r:1,c:7} },
    { s:{r:1,c:9}, e:{r:1,c:11} },
  ];
  wsIC["!rows"] = [
    {hpt:30},{hpt:20},{hpt:16},{hpt:24},
    ...icRows.map(()=>({hpt:32})),
    {hpt:20}
  ];
  wsIC["!cols"] = [
    {wch:10},{wch:32},{wch:10},{wch:18},{wch:18},
    {wch:18},{wch:36},{wch:16},{wch:8},{wch:8},
    {wch:12},{wch:16},{wch:16},{wch:12},{wch:10},
    {wch:12},{wch:12},{wch:12},{wch:14},{wch:14},
    {wch:24},{wch:10},
  ];

  // ── SHEET 2: Equipment List ────────────────────────────────
  const eqHdrs = [
    "Item No.","설비명","설비 유형","소속 Area",
    "재질","용량","설계 압력","설계 온도","연결 Interface 수","비고",
  ];
  const eqAoa = [
    [Object.assign(XS.hdr("Equipment List"), {
      s: XS.s({ bold:true, sz:12, bg:"1F3864", fc:"FFFFFF", bc:"1F3864" })
    }), ...Array(eqHdrs.length-1).fill(null)],
    [XS.c("작성일",{bold:true,bg:"D9E1F2",fc:"1F3864",h:"left"}),
     XS.c(toDate,  {bg:"FFFFFF",h:"left"}),
     null,
     XS.c("총 설비 수",{bold:true,bg:"D9E1F2",fc:"1F3864",h:"left"}),
     XS.c(`${equipRows.length}건`,{bold:true,bg:"FFFFFF",fc:"1D4ED8"}),
     ...Array(eqHdrs.length-5).fill(null)],
    eqHdrs.map(h => XS.shdr(h)),
    ...equipRows.map((row,i) => eqHdrs.map((h,j) => {
      const v = row[h] ?? "";
      if (h==="Item No.") return XS.alt(v, i, { bold:true, fc:"1D4ED8" });
      if (h==="연결 Interface 수") return XS.alt(
        v, i,
        { bold: v>0, fc: v>2?"DC2626":v>0?"EA580C":"16A34A" }
      );
      return XS.alt(v, i, { h: j<4?"center":"left" });
    })),
  ];
  const wsEq = aoaToSheet(eqAoa);
  wsEq["!merges"] = [
    { s:{r:0,c:0}, e:{r:0,c:eqHdrs.length-1} },
    { s:{r:1,c:1}, e:{r:1,c:2} },
  ];
  wsEq["!rows"] = [{hpt:28},{hpt:18},{hpt:22},...equipRows.map(()=>({hpt:22}))];
  wsEq["!cols"] = [
    {wch:12},{wch:22},{wch:20},{wch:22},
    {wch:12},{wch:14},{wch:12},{wch:12},{wch:16},{wch:28},
  ];

  // ── SHEET 3: Connection List ───────────────────────────────
  const cnHdrs = [
    "Line No.","Line Type","Fluid (Primary)","Fluid (Sub)",
    "Size","Schedule","Line Text",
    "From (Item No.)","To (Item No.)","연결 IC No.","IC 상태",
  ];
  const cnAoa = [
    [Object.assign(XS.hdr("Connection List"), {
      s: XS.s({ bold:true, sz:12, bg:"1F3864", fc:"FFFFFF", bc:"1F3864" })
    }), ...Array(cnHdrs.length-1).fill(null)],
    cnHdrs.map(h => XS.shdr(h)),
    ...connRows.map((row,i) => cnHdrs.map((h) => {
      const v = row[h] ?? "";
      if (h==="IC 상태") return XS.status(v||"—");
      if (h==="Line No."||h==="연결 IC No.")
        return XS.alt(v, i, { bold:true, fc:"1D4ED8" });
      if (h==="Fluid (Sub)") {
        const color = getFluidColor(v);
        const rgb = color.replace("#","");
        return XS.alt(v, i, { bold:true, fc: rgb });
      }
      return XS.alt(v, i);
    })),
  ];
  const wsCn = aoaToSheet(cnAoa);
  wsCn["!merges"] = [{ s:{r:0,c:0}, e:{r:0,c:cnHdrs.length-1} }];
  wsCn["!rows"] = [{hpt:28},{hpt:22},...connRows.map(()=>({hpt:22}))];
  wsCn["!cols"] = [
    {wch:12},{wch:12},{wch:14},{wch:10},
    {wch:8},{wch:8},{wch:16},
    {wch:16},{wch:16},{wch:12},{wch:12},
  ];

  // ── SHEET 4: Requirements ──────────────────────────────────
  const rqHdrs = [
    "Item No.","설비 유형","소속 Area","Stakeholder","날짜","요구사항",
  ];
  const rqAoa = [
    [Object.assign(XS.hdr("Requirements"), {
      s: XS.s({ bold:true, sz:12, bg:"1F3864", fc:"FFFFFF", bc:"1F3864" })
    }), ...Array(rqHdrs.length-1).fill(null)],
    rqHdrs.map(h => XS.shdr(h)),
    ...reqRows.map((row,i) => rqHdrs.map((h) => {
      const v = row[h] ?? "";
      const wrap = h==="요구사항";
      return XS.alt(v, i,
        { h: h==="요구사항"?"left":"center", wrap,
          bold: h==="Item No.", fc: h==="Item No."?"1D4ED8":"000000" });
    })),
    ...(reqRows.length===0 ? [[XS.c("등록된 요구사항이 없습니다.",
      {h:"center",fc:"94A3B8",border:false})]] : []),
  ];
  const wsRq = aoaToSheet(rqAoa);
  wsRq["!merges"] = [{ s:{r:0,c:0}, e:{r:0,c:rqHdrs.length-1} }];
  wsRq["!rows"] = [{hpt:28},{hpt:22},...reqRows.map(()=>({hpt:36}))];
  wsRq["!cols"] = [
    {wch:14},{wch:14},{wch:22},{wch:16},{wch:12},{wch:42},
  ];

  XLSX.utils.book_append_sheet(wb, wsIC,  "IC Register");
  XLSX.utils.book_append_sheet(wb, wsEq,  "Equipment List");
  XLSX.utils.book_append_sheet(wb, wsCn,  "Connection List");
  XLSX.utils.book_append_sheet(wb, wsRq,  "Requirements");

  XLSX.writeFile(wb, `MBSE_ICRegister_${today}.xlsx`);
  return { icCount: icRows.length, equipCount: equipRows.length, connCount: connRows.length };
};

// ─────────────────────────────────────────────────────────────
// EXCEL IMPORT — IC Register → MBSE 모델 반영
// IC Register 시트: 상태·담당자·비고 → Edge data 업데이트
// Equipment List 시트: 스펙 → Node data 업데이트
// Requirements 시트: 요구사항 병합
// ─────────────────────────────────────────────────────────────
const importFromExcel = async (file, nodes, edges, setNodes, setEdges) => {
  const XLSX = await loadXLSX();
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = ev => {
      try {
        const wb = XLSX.read(ev.target.result, { type:"array" });
        let updatedNodes = [...nodes];
        let updatedEdges = [...edges];
        const log = [];

        // ── 헬퍼: 스타일 적용 시트의 헤더 행 자동 감지 ───────
        // Export 시 타이틀(1)+프로젝트정보(1)+범례(1)+헤더(1) = 3행 스킵
        // "Edge ID" 또는 "Node ID" 컬럼이 있는 첫 행을 헤더로 사용
        const readSheet = (ws, keyCol) => {
          if (!ws) return [];
          // 먼저 raw로 읽어서 헤더 행 위치 찾기
          const raw = XLSX.utils.sheet_to_json(ws, { header:1 });
          let headerRow = 0;
          for (let i = 0; i < Math.min(raw.length, 10); i++) {
            if (raw[i] && raw[i].includes(keyCol)) {
              headerRow = i;
              break;
            }
          }
          // 헤더 행부터 파싱
          return XLSX.utils.sheet_to_json(ws, { range: headerRow });
        };

        // ── IC Register 시트 → Edge 업데이트 ─────────────
        const wsIC = wb.Sheets["IC Register"];
        if (wsIC) {
          const rows = readSheet(wsIC, "Edge ID");
          rows.forEach(row => {
            const edgeId = String(row["Edge ID"]||"").trim();
            if (!edgeId) return;
            const idx = updatedEdges.findIndex(e => e.id === edgeId);
            if (idx === -1) return;
            const e = updatedEdges[idx];
            updatedEdges[idx] = {
              ...e,
              data: {
                ...e.data,
                serialNo:    row["Line No."]      != null ? String(row["Line No."]||"")      : e.data?.serialNo,
                spec:        row["Schedule"]       != null ? String(row["Schedule"]||"")      : e.data?.spec,
                lineText:    row["Line Text"]      != null ? String(row["Line Text"]||"")     : e.data?.lineText,
                ic_no:       row["IC No."]         != null ? String(row["IC No."]||"")        : e.data?.ic_no,
                ic_status:   row["상태"]            != null ? String(row["상태"]||"")           : e.data?.ic_status,
                ic_priority: row["우선순위"]         != null ? String(row["우선순위"]||"")       : e.data?.ic_priority,
                ic_due:      row["목표 완료일"]      != null ? String(row["목표 완료일"]||"")    : e.data?.ic_due,
                ic_closed:   row["실제 완료일"]      != null ? String(row["실제 완료일"]||"")    : e.data?.ic_closed,
                ic_resp_from:row["담당자 (From)"]   != null ? String(row["담당자 (From)"]||"")  : e.data?.ic_resp_from,
                ic_resp_to:  row["담당자 (To)"]     != null ? String(row["담당자 (To)"]||"")    : e.data?.ic_resp_to,
                ic_remark:   row["비고"]            != null ? String(row["비고"]||"")           : e.data?.ic_remark,
                icd_no:      row["ICD 번호"]        != null ? String(row["ICD 번호"]||"")       : e.data?.icd_no,
              }
            };
            log.push(`IC ${row["IC No."]} → Edge ${edgeId} 업데이트`);
          });
        }

        // ── Equipment List 시트 → Node 업데이트 ──────────
        const wsEq = wb.Sheets["Equipment List"];
        if (wsEq) {
          const rows = readSheet(wsEq, "Node ID");
          rows.forEach(row => {
            const nodeId = String(row["Node ID"]||"").trim();
            if (!nodeId) return;
            const idx = updatedNodes.findIndex(n => n.id === nodeId);
            if (idx === -1) return;
            const n = updatedNodes[idx];
            updatedNodes[idx] = {
              ...n,
              data: {
                ...n.data,
                itemNo:   row["Item No."] != null ? String(row["Item No."]||"")  : n.data.itemNo,
                label:    row["설비명"]    != null ? String(row["설비명"]||"")     : n.data.label,
                material: row["재질"]      != null ? String(row["재질"]||"")      : n.data.material,
                capacity: row["용량"]      != null ? String(row["용량"]||"")      : n.data.capacity,
                designP:  row["설계 압력"] != null ? String(row["설계 압력"]||"") : n.data.designP,
                designT:  row["설계 온도"] != null ? String(row["설계 온도"]||"") : n.data.designT,
                summary:  row["비고"]      != null ? String(row["비고"]||"")      : n.data.summary,
              }
            };
            log.push(`Equipment ${row["Item No."]} → Node ${nodeId} 업데이트`);
          });
        }

        // ── Connection List 시트 → Edge 기본 속성 ────────
        const wsConn = wb.Sheets["Connection List"];
        if (wsConn) {
          const rows = readSheet(wsConn, "Edge ID");
          rows.forEach(row => {
            const edgeId = String(row["Edge ID"]||"").trim();
            if (!edgeId) return;
            const idx = updatedEdges.findIndex(e => e.id === edgeId);
            if (idx === -1) return;
            const e = updatedEdges[idx];
            const sizeRaw = row["Size"] != null ? String(row["Size"]||"") : "";
            const sizeNum = sizeRaw.replace(/[^0-9]/g,"");
            updatedEdges[idx] = {
              ...e,
              data: {
                ...e.data,
                serialNo:     row["Line No."]       != null ? String(row["Line No."]||"")    : e.data?.serialNo,
                lineType:     row["Line Type"]       != null ? String(row["Line Type"]||"")  : e.data?.lineType,
                fluidPrimary: row["Fluid (Primary)"] != null ? String(row["Fluid (Primary)"]||"") : e.data?.fluidPrimary,
                fluidSub:     row["Fluid (Sub)"]     != null ? String(row["Fluid (Sub)"]||"")     : e.data?.fluidSub,
                size:         sizeRaw,
                sizeNum:      sizeNum,
                spec:         row["Schedule"]        != null ? String(row["Schedule"]||"")   : e.data?.spec,
                lineText:     row["Line Text"]       != null ? String(row["Line Text"]||"")  : e.data?.lineText,
              }
            };
          });
        }

        // ── Requirements 시트 ─────────────────────────────
        const wsR = wb.Sheets["Requirements"];
        if (wsR) {
          const rows = readSheet(wsR, "Node ID");
          const reqMap = {};
          rows.forEach(r => {
            const nodeId = String(r["Node ID"]||"").trim();
            if (!nodeId || !r["요구사항"]) return;
            if (!reqMap[nodeId]) reqMap[nodeId] = [];
            reqMap[nodeId].push({
              id:   Date.now() + Math.random(),
              text: String(r["요구사항"]   || ""),
              who:  String(r["Stakeholder"]|| ""),
              date: String(r["날짜"]        || ""),
            });
          });
          updatedNodes = updatedNodes.map(n => {
            if (!reqMap[n.id]) return n;
            return { ...n, data:{ ...n.data, requirements: reqMap[n.id] } };
          });
        }

        setNodes(updatedNodes);
        setEdges(updatedEdges);
        resolve({ log, msg:`IC Register Import 완료 — ${log.length}건 반영` });
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
  /* ── 포트 핸들: 기본 숨김, hover/선택 시 표시 ── */
  .react-flow__handle {
    width: 10px !important;
    height: 10px !important;
    border-radius: 50% !important;
    background: #0d9488 !important;   /* teal */
    border: 2px solid #fff !important;
    opacity: 0 !important;
    transition: opacity 0.15s ease, transform 0.15s ease !important;
    z-index: 20 !important;
  }
  .react-flow__node:hover .react-flow__handle {
    opacity: 1 !important;
  }
  .react-flow__node.selected .react-flow__handle {
    opacity: 1 !important;
  }
  .react-flow__handle:hover {
    opacity: 1 !important;
    transform: scale(1.5) !important;
    background: #0f766e !important;
  }
  .react-flow__handle.connecting {
    opacity: 1 !important;
    transform: scale(1.3) !important;
  }
  /* 연결 드래그 중 전체 핸들 표시 */
  .react-flow__pane.connecting .react-flow__handle {
    opacity: 0.8 !important;
  }
  /* 범위 선택 박스 */
  .react-flow__selection {
    background: rgba(37,99,235,0.06) !important;
    border: 1.5px dashed #2563eb !important;
    border-radius: 4px !important;
  }
  .mbse-label-input {
    background: rgba(255,255,255,0.97);
    border: 1.5px solid #3b82f6;
    border-radius: 4px;
    padding: 2px 6px;
    font-size: 12px;
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
// Equipment 카테고리 분리
const EQUIPMENT_UT = [
  "Tank","Pump","Pond","Heat Exchanger","Filter","Hopper","Decanter",
  "Cooling Tower","Clarifier","Classifier","Feed Box","Chemical Dosing",
  "Gas Duct","Steel Structure",
];
const EQUIPMENT_ME = [
  "Reactor","Scrubber","Bag Filter","Feed Bin",
  "Bucket Elev.","Compressor","Fan","Stand Pipe","Bubbler",
  "Riser","Pnumatic Conv.","Conveyor","Machine","Structure","Hot Duct",
];
const EQUIPMENT_LIST = [...EQUIPMENT_UT, ...EQUIPMENT_ME];

// ── Connection 타입 ──────────────────────────────────────────
const CONNECTION_LIST = ["Piping","Piping (↔)","Duct","Brench","Process Gas","Material"];
const CONVEYOR_LIST   = ["Conveyor"];

// 라인 스타일 정의
const LINE_STYLE = {
  Piping:          { color:"#94a3b8", sw:2,   dash:"none", bidir:false },
  "Piping (↔)":   { color:"#94a3b8", sw:2,   dash:"none", bidir:true  },
  Duct:            { color:"#475569", sw:4,   dash:"none", bidir:false },
  Brench:          { color:"#94a3b8", sw:1.5, dash:"none", bidir:false },
  "Process Gas":   { color:"#7c3aed", sw:5,   dash:"none", bidir:false },
  Material:        { color:"#92400e", sw:5,   dash:"none", bidir:false },
  Conveyor:        { color:"#78350f", sw:2,   dash:"7,3",  bidir:false },
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

// EQUIP_DEFAULTS: 신규 spec 필드 키 기준 기본값
// (구버전 capacity/material/designP/designT 는 사용하지 않음)
const EQUIP_DEFAULTS = {
  Tank:             { medium:"Water",           capacity:"100",  designP:"5",   designT:"80",  material:"SS304"     },
  Pump:             { kindOfLiquid:"Water",     capacity:"50",   deliveryHead:"30", suctionPress:"0.1", dischargePress:"3.1", material:"SC410", impeller:"SSC13" },
  Pond:             { medium:"Process Water",   capacity:"500",  designP:"0.5", designT:"40",  material:"Concrete"  },
  "Heat Exchanger": { capacity:"500",           designP:"6",     designT:"80",  material:"SS316"     },
  Filter:           { flowMedium:"Water",       capacity:"20",   designP:"6",   designT:"80",  material:"CS"        },
  Hopper:           { medium:"Dry Solid",       capacity:"10",   designT:"50",  material:"MS"        },
  Decanter:         { feedLiquid:"Sludge",      capacity:"20",   designT:"60",  material:"SS304"     },
  "Cooling Tower":  { capacity:"1000",          inletTemp:"42",  outletTemp:"32", material:"FRP"     },
  Clarifier:        { feedLiquid:"Slurry",      capacity:"2200", designT:"70",  material:"CS+Epoxy"  },
  Classifier:       { feedLiquid:"Slurry",      capacity:"400",  designT:"60",  material:"SM45C"     },
  "Feed Box":       { feedLiquid:"Slurry",      capacity:"56",   operPress:"0", designT:"80",  material:"CS"        },
  "Chemical Dosing":{ kindOfLiquid:"Chemical",  capacity:"10",   deliveryHead:"50", material:"PP"    },
  Scrubber:         { medium:"Gas",             capacity:"10000",designP:"0.5", designT:"60",  material:"FRP"       },
  "Bag Filter":     { medium:"Gas",             capacity:"5000", designP:"0.3", designT:"180", material:"CS"        },
  Reactor:          { medium:"Gas",             capacity:"50",   designP:"10",  designT:"150", material:"SS316L"    },
  "Feed Bin":       { medium:"Solid",           capacity:"20",   designT:"50",  material:"MS"        },
  "Gas Duct":       { medium:"Gas",             capacity:"20000",designP:"0.5", designT:"300", material:"MS"        },
  "Steel Structure":{ material:"A36 Steel"                                                            },
  "Bucket Elev.":   { medium:"Solid",           capacity:"50",   material:"MS"                        },
  Compressor:       { medium:"Gas",             capacity:"5000", inletPress:"11",dischargePress:"38", material:"CS" },
  Fan:              { medium:"Air",             capacity:"20000",staticPress:"450", material:"CS"     },
  "Stand Pipe":     { medium:"Gas",             capacity:"10",   designP:"5",   designT:"300", material:"MS"        },
  Bubbler:          { medium:"Gas",             designP:"2",     designT:"200", material:"MS"         },
  Riser:            { medium:"Gas",             designP:"3",     designT:"300", material:"MS"         },
  "Pnumatic Conv.": { medium:"Solid",           capacity:"30",   material:"MS"                        },
  Conveyor:         { medium:"Solid",           capacity:"100",  material:"MS"                        },
  Machine:          { material:"MS"                                                                    },
  Structure:        { material:"A36 Steel"                                                             },
  "Hot Duct":       { medium:"Hot Gas",         capacity:"30000",designP:"0.3", designT:"900", material:"MS"       },
};

// ─────────────────────────────────────────────────────────────
// 설비 유형별 맞춤 사양 항목 정의
// ─────────────────────────────────────────────────────────────
const EQUIP_SPEC_FIELDS = {
  // Pump 계열
  Pump: [
    { key:"kindOfLiquid",   label:"Kind of Liquid",        unit:""        },
    { key:"capacity",       label:"Capacity",              unit:"m³/h"    },
    { key:"deliveryHead",   label:"Delivery Head",         unit:"m"       },
    { key:"suctionPress",   label:"Suction Press.",        unit:"barg"    },
    { key:"dischargePress", label:"Discharge Press.",      unit:"barg"    },
    { key:"liquidTemp",     label:"Liquid Temp.",          unit:"℃"       },
    { key:"construction",   label:"Type of Construction",  unit:""        },
    { key:"shaftSeal",      label:"Shaft Seal",            unit:""        },
    { key:"material",       label:"Casing Material",       unit:""        },
    { key:"impeller",       label:"Impeller Material",     unit:""        },
  ],
  // Heat Exchanger 계열
  "Heat Exchanger": [
    { key:"flowRatePri",    label:"Flow Rate (Primary)",   unit:"m³/h"    },
    { key:"flowRateSec",    label:"Flow Rate (Secondary)", unit:"m³/h"    },
    { key:"inletTempPri",   label:"Inlet Temp. (Pri.)",    unit:"℃"       },
    { key:"outletTempPri",  label:"Outlet Temp. (Pri.)",   unit:"℃"       },
    { key:"inletTempSec",   label:"Inlet Temp. (Sec.)",    unit:"℃"       },
    { key:"outletTempSec",  label:"Outlet Temp. (Sec.)",   unit:"℃"       },
    { key:"designPress",    label:"Design Press.",         unit:"barg"    },
    { key:"designTemp",     label:"Design Temp.",          unit:"℃"       },
    { key:"capacity",       label:"Heat Duty",             unit:"kW"      },
    { key:"material",       label:"Plate/Tube Material",   unit:""        },
  ],
  // Tank / Vessel 계열
  Tank: [
    { key:"medium",         label:"Medium",                unit:""        },
    { key:"capacity",       label:"Effective Volume",      unit:"m³"      },
    { key:"designP",        label:"Design Press.",         unit:"barg"    },
    { key:"designT",        label:"Design Temp.",          unit:"℃"       },
    { key:"operPress",      label:"Operating Press.",      unit:"barg"    },
    { key:"operTemp",       label:"Operating Temp.",       unit:"℃"       },
    { key:"material",       label:"Shell Material",        unit:""        },
    { key:"installation",   label:"Installation",          unit:""        },
  ],
  // Cooling Tower
  "Cooling Tower": [
    { key:"capacity",       label:"Total Capacity",        unit:"m³/h"    },
    { key:"inletTemp",      label:"Inlet Temp.",           unit:"℃"       },
    { key:"outletTemp",     label:"Outlet Temp.",          unit:"℃"       },
    { key:"wetBulbTemp",    label:"Wet Bulb Temp.",        unit:"℃"       },
    { key:"numCells",       label:"Number of Cells",       unit:"pcs"     },
    { key:"construction",   label:"Type of Construction",  unit:""        },
    { key:"material",       label:"Filling Material",      unit:""        },
  ],
  // Filter
  Filter: [
    { key:"flowMedium",     label:"Flow Medium",           unit:""        },
    { key:"capacity",       label:"Flow Rate",             unit:"m³/h"    },
    { key:"designP",        label:"Design Press.",         unit:"barg"    },
    { key:"designT",        label:"Design Temp.",          unit:"℃"       },
    { key:"meshSize",       label:"Mesh Size",             unit:"μm"      },
    { key:"pressureDrop",   label:"Pressure Drop",         unit:"barg"    },
    { key:"material",       label:"Casing Material",       unit:""        },
    { key:"filterType",     label:"Filter Type",           unit:""        },
  ],
  // Clarifier
  Clarifier: [
    { key:"feedLiquid",     label:"Feed Liquid",           unit:""        },
    { key:"capacity",       label:"Flow Rate",             unit:"m³/h"    },
    { key:"designT",        label:"Design Temp.",          unit:"℃"       },
    { key:"overflowSolid",  label:"Overflow Solid",        unit:"mg/l"    },
    { key:"underflowSolid", label:"Underflow Solid",       unit:"g/l"     },
    { key:"clarifierSize",  label:"Clarifier Size",        unit:""        },
    { key:"construction",   label:"Type of Construction",  unit:""        },
    { key:"material",       label:"Frame Material",        unit:""        },
  ],
  // Decanter
  Decanter: [
    { key:"feedLiquid",     label:"Feed Liquid",           unit:""        },
    { key:"capacity",       label:"Capacity",              unit:"m³/h"    },
    { key:"solidContent",   label:"Solid Content",         unit:"g/l"     },
    { key:"residualMoisture",label:"Residual Moisture",    unit:"%"       },
    { key:"material",       label:"Bowl Material",         unit:""        },
    { key:"scroll",         label:"Scroll Material",       unit:""        },
  ],
  // Compressor / Fan
  Compressor: [
    { key:"medium",         label:"Medium",                unit:""        },
    { key:"capacity",       label:"Volume Flow Rate",      unit:"Nm³/h"   },
    { key:"inletPress",     label:"Inlet Press.",          unit:"barg"    },
    { key:"dischargePress", label:"Discharge Press.",      unit:"barg"    },
    { key:"inletTemp",      label:"Inlet Temp.",           unit:"℃"       },
    { key:"outletTemp",     label:"Outlet Temp.",          unit:"℃"       },
    { key:"construction",   label:"Type of Design",        unit:""        },
    { key:"numStages",      label:"Number of Stages",      unit:""        },
  ],
  Fan: [
    { key:"medium",         label:"Medium",                unit:""        },
    { key:"capacity",       label:"Volume Flow Rate",      unit:"m³/h"    },
    { key:"staticPress",    label:"Static Press.",         unit:"mbar"    },
    { key:"inletTemp",      label:"Inlet Temp.",           unit:"℃"       },
    { key:"noiseLevel",     label:"Noise Level",           unit:"dB(A)"   },
    { key:"material",       label:"Casing/Impeller",       unit:""        },
  ],
  // Chemical Dosing
  "Chemical Dosing": [
    { key:"kindOfLiquid",   label:"Kind of Liquid",        unit:""        },
    { key:"capacity",       label:"Pump Capacity",         unit:"l/h"     },
    { key:"deliveryHead",   label:"Delivery Head",         unit:"m"       },
    { key:"tankVolume",     label:"Tank Volume",           unit:"m³"      },
    { key:"liquidTemp",     label:"Liquid Temp.",          unit:"℃"       },
    { key:"material",       label:"Tank Material",         unit:""        },
    { key:"pumpType",       label:"Type of Pump",          unit:""        },
  ],
  // Classifier
  Classifier: [
    { key:"feedLiquid",     label:"Feed Liquid",           unit:""        },
    { key:"capacity",       label:"Flow Rate",             unit:"m³/h"    },
    { key:"solidsFlow",     label:"Solids Flow",           unit:"t/h"     },
    { key:"particleSize",   label:"Particle Size",         unit:"mm"      },
    { key:"designT",        label:"Design Temp.",          unit:"℃"       },
    { key:"construction",   label:"Type of Construction",  unit:""        },
    { key:"material",       label:"Shaft Material",        unit:""        },
  ],
  // Pond
  Pond: [
    { key:"medium",         label:"Medium",                unit:""        },
    { key:"capacity",       label:"Effective Volume",      unit:"m³"      },
    { key:"designP",        label:"Design Press.",         unit:"barg"    },
    { key:"designT",        label:"Design Temp.",          unit:"℃"       },
    { key:"material",       label:"Material",              unit:""        },
  ],
  // Feed Box
  "Feed Box": [
    { key:"feedLiquid",     label:"Feed Liquid",           unit:""        },
    { key:"capacity",       label:"Volume",                unit:"m³"      },
    { key:"operPress",      label:"Operating Press.",      unit:"barg"    },
    { key:"designT",        label:"Design Temp.",          unit:"℃"       },
    { key:"material",       label:"Shell Material",        unit:""        },
  ],
  // Hopper / Feed Bin / Cake Hopper
  Hopper: [
    { key:"medium",         label:"Medium",                unit:""        },
    { key:"capacity",       label:"Effective Volume",      unit:"m³"      },
    { key:"designT",        label:"Design Temp.",          unit:"℃"       },
    { key:"material",       label:"Shell Material",        unit:""        },
  ],
  "Feed Bin": [
    { key:"medium",         label:"Medium",                unit:""        },
    { key:"capacity",       label:"Effective Volume",      unit:"m³"      },
    { key:"designT",        label:"Design Temp.",          unit:"℃"       },
    { key:"material",       label:"Shell Material",        unit:""        },
  ],
  // Scrubber / Reactor / Bag Filter
  Scrubber: [
    { key:"medium",         label:"Medium",                unit:""        },
    { key:"capacity",       label:"Volume Flow Rate",      unit:"Nm³/h"   },
    { key:"designP",        label:"Design Press.",         unit:"barg"    },
    { key:"designT",        label:"Design Temp.",          unit:"℃"       },
    { key:"material",       label:"Casing Material",       unit:""        },
  ],
  Reactor: [
    { key:"medium",         label:"Medium",                unit:""        },
    { key:"capacity",       label:"Effective Volume",      unit:"m³"      },
    { key:"designP",        label:"Design Press.",         unit:"barg"    },
    { key:"designT",        label:"Design Temp.",          unit:"℃"       },
    { key:"operPress",      label:"Operating Press.",      unit:"barg"    },
    { key:"operTemp",       label:"Operating Temp.",       unit:"℃"       },
    { key:"material",       label:"Shell Material",        unit:""        },
  ],
  "Bag Filter": [
    { key:"medium",         label:"Medium",                unit:""        },
    { key:"capacity",       label:"Volume Flow Rate",      unit:"Nm³/h"   },
    { key:"designP",        label:"Design Press.",         unit:"barg"    },
    { key:"designT",        label:"Design Temp.",          unit:"℃"       },
    { key:"material",       label:"Casing Material",       unit:""        },
  ],
};

// 기본 공통 사양 (유형별 정의가 없을 때)
const DEFAULT_SPEC_FIELDS = [
  { key:"capacity",  label:"Capacity",        unit:""    },
  { key:"designP",   label:"Design Press.",   unit:"barg"},
  { key:"designT",   label:"Design Temp.",    unit:"℃"   },
  { key:"material",  label:"Material",        unit:""    },
];

// 설비 유형에 맞는 사양 필드 반환
const getSpecFields = (equipType) =>
  EQUIP_SPEC_FIELDS[equipType] || DEFAULT_SPEC_FIELDS;
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
// Brench(분기점)는 무시하고 실제 근원지(Equipment/Instrument)를 추적
// ─────────────────────────────────────────────────────────────

// Brench를 거슬러 올라가 실제 근원 노드 찾기
const traceSource = (nodeId, allNodes, allEdges, visited=new Set()) => {
  if (visited.has(nodeId)) return nodeId;
  visited.add(nodeId);
  const node = allNodes.find(n => n.id === nodeId);
  if (!node) return nodeId;
  if (node.type !== "brench") return nodeId; // Equipment/Area → 여기가 진짜 근원
  // Brench → 이 노드로 들어오는 엣지의 source를 재귀 추적
  const inEdge = allEdges.find(e => e.target === nodeId);
  if (!inEdge) return nodeId;
  return traceSource(inEdge.source, allNodes, allEdges, visited);
};

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
  const seen = new Set(); // 중복 방지

  allEdges.forEach(e=>{
    const si=insideIds.has(e.source),ti=insideIds.has(e.target);
    if(si===ti) return; // 경계 통과 아님

    const d=e.data||{};
    const sub=d.fluidSub||d.lineType||"—";
    const size=d.size?` ${d.size}`:"";
    const label=`${sub}${size}`;

    if(ti){
      // 외부에서 안으로 들어옴 — source 추적 (Brench 스킵)
      const realSrcId = traceSource(e.source, allNodes, allEdges);
      const realSrc = allNodes.find(n=>n.id===realSrcId);
      // Brench면 근원지 이름 대신 라인 라벨 사용
      const srcName = (realSrc && realSrc.type!=="brench")
        ? (realSrc.data?.itemNo||realSrc.data?.label||"외부")
        : "외부";
      const key = `IN:${label}:${srcName}`;
      if(!seen.has(key)){ seen.add(key); inlets.push(label); }
    } else {
      // 안에서 밖으로 나감 — target 추적
      const realTgtId = traceSource(e.target, allNodes, allEdges);
      const realTgt = allNodes.find(n=>n.id===realTgtId);
      const tgtName = (realTgt && realTgt.type!=="brench")
        ? (realTgt.data?.itemNo||realTgt.data?.label||"외부")
        : "외부";
      const key = `OUT:${label}:${tgtName}`;
      if(!seen.has(key)){ seen.add(key); outlets.push(label); }
    }
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
      <NodeResizer
        minWidth={180} minHeight={120}
        isVisible={selected}
        lineStyle={{ border:"2px dashed #f59e0b" }}
        handleStyle={{ width:14, height:14, background:"#f59e0b", border:"2px solid #fff", borderRadius:3 }}
      />

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
            <span style={{
              fontSize: data.areaType==="Plant" ? 30
                      : data.areaType==="System"  ? 15
                      : data.areaType==="Package" ? 15
                      : 20,  // Item
              fontWeight:800, color:c.label, lineHeight:1.2,
            }}>
              {/* Plant/System/Package 는 태그 미표시, Item만 표시 */}
              {data.areaType==="Item" ? `[Item] ` : ""}
              {data.label || <span style={{ opacity:0.4, fontSize:"0.7em" }}>(더블클릭 편집)</span>}
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
      padding:"6px 12px 8px",fontSize:11,cursor:"default",
      position:"relative",userSelect:"none",boxSizing:"border-box",
    }}>
      <NodeResizer minWidth={80} minHeight={50} isVisible={selected} handleStyle={{ width:8,height:8 }}/>
      {handleList.map(({dir,pid,pct})=>(
        <Handle key={pid} type="source" position={posMap[dir]} id={pid} style={getStyle(dir,pct)}/>
      ))}

      {/* ── Item No: 상단 표시 ── */}
      {editing ? (
        <input ref={inputRef} className="mbse-label-input" value={draft}
          onChange={e=>setDraft(e.target.value)} onBlur={commitEdit}
          onKeyDown={e=>{ if(e.key==="Enter") commitEdit(); if(e.key==="Escape") setEditing(false); }}
          onClick={e=>e.stopPropagation()} placeholder="Item No"
          style={{ marginBottom:4 }}/>
      ) : (
        <div onDoubleClick={startEdit}
          style={{ fontWeight:700,color:"#1d4ed8",fontSize:10,textAlign:"center",
                   lineHeight:1.2,cursor:"text",padding:"1px 4px",borderRadius:3,
                   minWidth:60,marginBottom:3,
                   background: data.itemNo?"rgba(29,78,216,0.06)":"transparent",
                   border: data.itemNo?"1px dashed #bfdbfe":"1px dashed transparent" }}
          title="더블클릭으로 Item No 편집">
          {data.itemNo || <span style={{ color:"#cbd5e1",fontSize:9 }}>Item No</span>}
        </div>
      )}

      {/* ── 아이콘 ── */}
      <div style={{ color:"#3b82f6",marginBottom:3 }}><EquipSVG type={data.equipType} size={24}/></div>

      {/* ── 설비명 ── */}
      <div style={{ fontWeight:600,color:"#475569",fontSize:10,textAlign:"center",lineHeight:1.2 }}>
        {data.label && data.label!==(data.itemNo||data.equipType)
          ? data.label
          : data.equipType}
      </div>
    </div>
  );
});

// ─────────────────────────────────────────────────────────────
// BRENCH NODE
// ─────────────────────────────────────────────────────────────
const BrenchNode = memo(({ id, data, selected }) => (
  <div style={{
    width:24, height:24, borderRadius:"50%",
    background: selected ? "#f59e0b" : "#334155",
    border:`2.5px solid ${selected?"#b45309":"#94a3b8"}`,
    position:"relative", cursor:"move",
    boxShadow: selected ? "0 0 0 3px rgba(245,158,11,0.3)" : "0 2px 6px rgba(0,0,0,0.25)",
    display:"flex", alignItems:"center", justifyContent:"center",
  }}>
    {/* 중앙 십자 표시 */}
    <div style={{ width:10,height:2,background:"rgba(255,255,255,0.7)",position:"absolute" }}/>
    <div style={{ width:2,height:10,background:"rgba(255,255,255,0.7)",position:"absolute" }}/>
    {DIRS.map(pos=>(
      <Handle key={pos} type="source" position={pos} id={DIR_ID[pos]}
        style={{ width:8,height:8,borderRadius:"50%",background:"#0d9488",border:"2px solid #fff",zIndex:10 }}/>
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
// ─────────────────────────────────────────────────────────────
// ORTHOGONAL PATH BUILDER
// waypoints 기반 직각 꺾임 경로 (수평→수직)
// ─────────────────────────────────────────────────────────────
// ═══════════════════════════════════════════════════════════════
// PIPE EDGE SYSTEM — 완전 재작성
//
// 발췌 적용 (첨부 프롬프트):
//  ① 포트방향 인식 orthogonal 자동 라우팅 (노드 bbox 회피)
//  ② 꺾임 radius=6px (Q 베지어)
//  ③ 화살표: 채워진 삼각형 (10×6px)
//  ④ 선택 시 세그먼트 중앙 핸들 (수평↕ / 수직↔ 파란 사각형)
//  ⑤ 라벨: 가장 긴 세그먼트 중앙
// ═══════════════════════════════════════════════════════════════

const ELBOW_R = 4;   // 꺾임 반경 줄임 (기존 6→4)
const STUB    = 16;  // 포트 이탈 거리 줄임 (기존 30→16)

// ── 포트 방향 → 이탈 벡터 ────────────────────────────────────
const portVec = pos => {
  if (pos === "right")  return { dx: 1, dy: 0 };
  if (pos === "left")   return { dx:-1, dy: 0 };
  if (pos === "bottom") return { dx: 0, dy: 1 };
  if (pos === "top")    return { dx: 0, dy:-1 };
  return { dx: 1, dy: 0 };
};

// ── 경로 정규화 ───────────────────────────────────────────────
// 두 가지 처리:
//  1) 스냅: 거의 같은 x 또는 y 좌표(≤ SNAP_TOL)를 정확히 일치시킴
//  2) 병합: 연속된 같은 방향 세그먼트의 중간 waypoint 제거
const SNAP_TOL = 8; // px — 이 이내면 동일 좌표로 정렬

const normalizePath = (pts) => {
  if (!pts || pts.length < 2) return pts;

  // Step 1: 이웃한 포인트끼리 x/y 오차가 SNAP_TOL 이하면 정렬
  // 직전 세그먼트의 방향에 따라 수직 좌표만 맞춤
  const snapped = pts.map((p, i) => ({ ...p }));
  for (let i = 1; i < snapped.length; i++) {
    const prev = snapped[i - 1];
    const curr = snapped[i];
    // 수평으로 이동하려는 세그먼트: y가 거의 같으면 prev.y로 맞춤
    if (Math.abs(curr.y - prev.y) <= SNAP_TOL && Math.abs(curr.x - prev.x) > SNAP_TOL) {
      snapped[i] = { x: curr.x, y: prev.y };
    }
    // 수직으로 이동하려는 세그먼트: x가 거의 같으면 prev.x로 맞춤
    if (Math.abs(curr.x - prev.x) <= SNAP_TOL && Math.abs(curr.y - prev.y) > SNAP_TOL) {
      snapped[i] = { x: prev.x, y: curr.y };
    }
  }

  // Step 2: 완전히 중복된 점 제거
  const deduped = snapped.filter((p, i) =>
    i === 0 || Math.hypot(p.x - snapped[i-1].x, p.y - snapped[i-1].y) > 0.5
  );

  // Step 3: 세 점이 일직선(같은 x 또는 같은 y 연속)이면 중간 점 제거
  // 예) (100,50) → (100,50) → (200,50) : 세 점이 수평 → 중간 제거
  let merged = [deduped[0]];
  for (let i = 1; i < deduped.length - 1; i++) {
    const prev = merged[merged.length - 1];
    const curr = deduped[i];
    const next = deduped[i + 1];
    // prev→curr와 curr→next가 같은 방향이면 curr는 불필요
    const sameDirX = Math.abs(prev.y - curr.y) < 0.5 && Math.abs(curr.y - next.y) < 0.5;
    const sameDirY = Math.abs(prev.x - curr.x) < 0.5 && Math.abs(curr.x - next.x) < 0.5;
    if (!sameDirX && !sameDirY) {
      merged.push(curr);
    }
  }
  merged.push(deduped[deduped.length - 1]);

  return merged;
};

// ── 포트방향 인식 직각 라우팅 (개선) ─────────────────────────
// 핵심 개선:
//  ① stub 길이를 거리에 비례해서 동적 계산 → 교차 방지
//  ② 반대 방향 포트(left↔right, top↔bottom) 직결 최적화
//  ③ 같은 방향 포트(right→right 등) U턴 처리
const routePipe = (sx, sy, srcPos, tx, ty, tgtPos, storedWp) => {
  if (storedWp && storedWp.length > 0) {
    return normalizePath([{ x:sx,y:sy }, ...storedWp, { x:tx,y:ty }]);
  }

  const sv   = portVec(srcPos);
  const tv   = portVec(tgtPos);

  // 두 포트 간 거리에 비례한 stub 계산
  const dist = Math.hypot(tx-sx, ty-sy);
  const stub = Math.max(8, Math.min(STUB, dist * 0.18));

  const sx2  = sx + sv.dx * stub;
  const sy2  = sy + sv.dy * stub;
  const tx2  = tx + tv.dx * stub;
  const ty2  = ty + tv.dy * stub;

  let raw;

  // ── 케이스 1: 수평 출발·수평 도착 (right→left / left→right) ──
  if (sv.dy === 0 && tv.dy === 0) {
    if (Math.abs(sy - ty) < 2) {
      // 같은 y → 단순 수평 직선
      raw = [{ x:sx,y:sy }, { x:tx,y:ty }];
    } else if ((sv.dx > 0 && tx > sx) || (sv.dx < 0 && tx < sx)) {
      // 반대 방향 → 가운데에서 수직 연결
      const mx2 = (sx2 + tx2) / 2;
      raw = [
        { x:sx,  y:sy  },
        { x:mx2, y:sy  },
        { x:mx2, y:ty  },
        { x:tx,  y:ty  },
      ];
    } else {
      // U턴 (같은 방향 포트)
      const ux = sv.dx > 0
        ? Math.max(sx2, tx2) + stub
        : Math.min(sx2, tx2) - stub;
      raw = [
        { x:sx,  y:sy  },
        { x:ux,  y:sy  },
        { x:ux,  y:ty  },
        { x:tx,  y:ty  },
      ];
    }
  }
  // ── 케이스 2: 수직 출발·수직 도착 (bottom→top / top→bottom) ──
  else if (sv.dx === 0 && tv.dx === 0) {
    if (Math.abs(sx - tx) < 2) {
      raw = [{ x:sx,y:sy }, { x:tx,y:ty }];
    } else if ((sv.dy > 0 && ty > sy) || (sv.dy < 0 && ty < sy)) {
      const my2 = (sy2 + ty2) / 2;
      raw = [
        { x:sx,  y:sy  },
        { x:sx,  y:my2 },
        { x:tx,  y:my2 },
        { x:tx,  y:ty  },
      ];
    } else {
      const uy = sv.dy > 0
        ? Math.max(sy2, ty2) + stub
        : Math.min(sy2, ty2) - stub;
      raw = [
        { x:sx,  y:sy  },
        { x:sx,  y:uy  },
        { x:tx,  y:uy  },
        { x:tx,  y:ty  },
      ];
    }
  }
  // ── 케이스 3: 수평 출발·수직 도착 (right/left → top/bottom) ──
  else if (sv.dy === 0 && tv.dx === 0) {
    // src에서 수평으로 tx까지, 그 다음 수직으로 ty까지
    raw = [
      { x:sx,  y:sy  },
      { x:tx,  y:sy  },
      { x:tx,  y:ty  },
    ];
  }
  // ── 케이스 4: 수직 출발·수평 도착 (top/bottom → right/left) ──
  else {
    // src에서 수직으로 ty까지, 그 다음 수평으로 tx까지
    raw = [
      { x:sx,  y:sy  },
      { x:sx,  y:ty  },
      { x:tx,  y:ty  },
    ];
  }

  return normalizePath(raw);
};

// ── Rounded Elbow SVG path ────────────────────────────────────
// 중복점 제거 → 각 꺾임점에 Q 베지어 적용
const buildElbowPath = (pts, r = ELBOW_R) => {
  if (!pts || pts.length < 2) return "";

  const clean = pts.filter((p, i) =>
    i === 0 || Math.hypot(p.x - pts[i-1].x, p.y - pts[i-1].y) > 0.5
  );
  if (clean.length < 2) return "";

  let d = `M ${clean[0].x} ${clean[0].y}`;
  for (let i = 1; i < clean.length; i++) {
    const prev = clean[i-1], curr = clean[i], next = clean[i+1];
    if (!next) {
      d += ` L ${curr.x} ${curr.y}`;
    } else {
      const dx1 = curr.x-prev.x, dy1 = curr.y-prev.y;
      const dx2 = next.x-curr.x, dy2 = next.y-curr.y;
      const l1  = Math.hypot(dx1,dy1)||1, l2 = Math.hypot(dx2,dy2)||1;
      const rr  = Math.min(r, l1/2, l2/2);
      d += ` L ${curr.x-(dx1/l1)*rr} ${curr.y-(dy1/l1)*rr}`;
      d += ` Q ${curr.x} ${curr.y}`;
      d += ` ${curr.x+(dx2/l2)*rr} ${curr.y+(dy2/l2)*rr}`;
    }
  }
  return d;
};

// ── 세그먼트 목록 추출 ──────────────────────────────────────
// [{x1,y1,x2,y2, isHoriz, mx,my}]
const getSegments = pts => {
  const segs = [];
  for (let i = 0; i < pts.length-1; i++) {
    const a = pts[i], b = pts[i+1];
    const len = Math.hypot(b.x-a.x, b.y-a.y);
    segs.push({
      x1:a.x, y1:a.y, x2:b.x, y2:b.y,
      isHoriz: Math.abs(a.y-b.y) < 2,
      mx:(a.x+b.x)/2, my:(a.y+b.y)/2,
      len,
    });
  }
  return segs;
};

// ── PipeEdge ─────────────────────────────────────────────────
const PipeEdge = ({
  id,
  sourceX, sourceY, sourcePosition,
  targetX, targetY, targetPosition,
  source, target,
  data, selected,
}) => {
  // getNodes로 모든 노드 bbox 참조 (장애물 회피용)
  const { getNodes } = useReactFlow();

  // 라인 스타일
  const lt        = data?.lineType || "Piping";
  const ls        = LINE_STYLE[lt] || LINE_STYLE.Piping;
  const baseColor = data?.fluidSub ? getFluidColor(data.fluidSub) : ls.color;
  const stroke    = selected ? "#f59e0b" : baseColor;
  const sw        = ls.sw || 2;
  const mkId      = `arrow_${id}`;

  // IC / 라벨 데이터
  const icStatusColor = {
    "OPEN":"#CA8A04","IN PROGRESS":"#2563EB",
    "CLOSED":"#16A34A","OVERDUE":"#DC2626",
  };
  const fluidLabel  = data?.fluidSub || "";
  const sizeLabel   = data?.sizeNum ? `${data.sizeNum}A` : (data?.size||"");
  const pipingLabel = [fluidLabel,sizeLabel].filter(Boolean).join("-");
  const isSpecial   = lt==="Process Gas"||lt==="Material";
  const showLabel   = isSpecial?(data?.lineText||lt):pipingLabel;
  const icNo        = data?.ic_no    || "";
  const icStatus    = data?.ic_status|| "";
  const icColor     = icStatusColor[icStatus]||"#64748B";

  // ── 장애물(Equipment) bbox 수집 ─────────────────────────
  // Area는 통과 허용, equipment/instrument/brench는 회피
  const allNodes = getNodes();
  const obstacles = allNodes
    .filter(n => n.id !== source && n.id !== target
              && n.type !== "area")
    .map(n => ({
      x1: (n.position?.x||0) - 10,
      y1: (n.position?.y||0) - 10,
      x2: (n.position?.x||0) + (n.width||120) + 10,
      y2: (n.position?.y||0) + (n.height||60) + 10,
    }));

  // ── 선분이 장애물과 교차하는지 검사 ──────────────────────
  const segIntersectsBox = (x1,y1,x2,y2, box) => {
    // 선분 bbox와 box 겹침 확인 (빠른 판별)
    const minX=Math.min(x1,x2), maxX=Math.max(x1,x2);
    const minY=Math.min(y1,y2), maxY=Math.max(y1,y2);
    return !(maxX < box.x1 || minX > box.x2 || maxY < box.y1 || minY > box.y2);
  };

  // ── 장애물 회피 경로 보정 ─────────────────────────────────
  // 자동 라우팅 결과의 각 세그먼트를 검사하여
  // 충돌 시 우회 경유점을 삽입
  const avoidObstacles = (pts) => {
    if (obstacles.length === 0) return pts;
    let result = [...pts];
    let maxIter = 4; // 최대 반복 횟수
    while (maxIter-- > 0) {
      let changed = false;
      const next = [result[0]];
      for (let i = 0; i < result.length - 1; i++) {
        const a = result[i], b = result[i+1];
        const isH = Math.abs(a.y - b.y) < 2; // 수평 세그먼트
        let hit = null;
        for (const box of obstacles) {
          if (segIntersectsBox(a.x, a.y, b.x, b.y, box)) {
            hit = box; break;
          }
        }
        if (hit) {
          changed = true;
          if (isH) {
            // 수평 세그먼트가 박스와 충돌 → 위 또는 아래로 우회
            const margin = 20;
            // 위/아래 중 더 여유있는 방향
            const aboveY = hit.y1 - margin;
            const belowY = hit.y2 + margin;
            const midX   = (a.x + b.x) / 2;
            const byY    = Math.abs(a.y - aboveY) < Math.abs(a.y - belowY)
                           ? aboveY : belowY;
            next.push({ x:a.x,  y:byY },
                       { x:b.x,  y:byY });
          } else {
            // 수직 세그먼트가 박스와 충돌 → 왼쪽 또는 오른쪽 우회
            const margin = 20;
            const leftX  = hit.x1 - margin;
            const rightX = hit.x2 + margin;
            const byX    = Math.abs(a.x - leftX) < Math.abs(a.x - rightX)
                           ? leftX : rightX;
            next.push({ x:byX, y:a.y },
                       { x:byX, y:b.y });
          }
        } else {
          next.push(b);
        }
      }
      if (!changed) break;
      result = normalizePath(next);
    }
    return result;
  };

  // ── 경로 계산 ────────────────────────────────────────────
  const storedWp = data?.waypoints||[];
  const rawPts = routePipe(
    sourceX, sourceY, sourcePosition||"right",
    targetX, targetY, targetPosition||"left",
    storedWp
  );
  // storedWp 없을 때만 장애물 회피 적용
  const pts  = storedWp.length > 0 ? rawPts : avoidObstacles(rawPts);
  const path = buildElbowPath(pts);
  const segs = getSegments(pts);

  // 전체 라인 길이 계산 → 화살표 크기 동적 조정
  const totalLen = segs.reduce((s,g)=>s+g.len, 0);
  // 짧은 선일수록 화살표 작게 (최소 4, 최대 8, strokeWidth 기준)
  const arrowScale = totalLen < 40 ? 4 : totalLen < 80 ? 5 : 7;

  // 라벨 위치: 가장 긴 세그먼트 중앙
  const longest = segs.reduce((a,b)=>b.len>a.len?b:a, segs[0]||{mx:0,my:0,len:0});

  // SVG 좌표 변환
  const toSVG = (svg, ev) => {
    const pt = svg.createSVGPoint();
    pt.x=ev.clientX; pt.y=ev.clientY;
    return pt.matrixTransform(svg.getScreenCTM().inverse());
  };

  // ── 세그먼트 핸들 드래그 ───────────────────────────────
  // 수평 세그먼트: y 이동 (↕)
  // 수직 세그먼트: x 이동 (↔)
  const onSegDrag = useCallback((e, segIdx) => {
    e.stopPropagation();
    const svg = e.target.closest("svg"); if(!svg) return;
    const seg      = segs[segIdx];
    const origPos  = toSVG(svg, e);
    // wps 초기화: 현재 pts에서 src·tgt 제외
    const initWps  = storedWp.length>0
      ? storedWp.map(p=>({...p}))
      : pts.slice(1,-1).map(p=>({...p}));

    const onMove = mv => {
      const cur = toSVG(svg, mv);
      const nw  = initWps.map(p=>({...p}));
      const wi1 = segIdx-1, wi2 = segIdx;
      if (seg.isHoriz) {
        const dy = cur.y - origPos.y;
        if (wi1>=0 && wi1<nw.length) nw[wi1]={...nw[wi1], y:initWps[wi1].y+dy};
        if (wi2>=0 && wi2<nw.length) nw[wi2]={...nw[wi2], y:(initWps[wi2]?.y??cur.y)+dy};
      } else {
        const dx = cur.x - origPos.x;
        if (wi1>=0 && wi1<nw.length) nw[wi1]={...nw[wi1], x:initWps[wi1].x+dx};
        if (wi2>=0 && wi2<nw.length) nw[wi2]={...nw[wi2], x:(initWps[wi2]?.x??cur.x)+dx};
      }
      // 정규화: 거의 일직선인 꺾임 제거
      const allPts  = normalizePath([{x:sourceX,y:sourceY},...nw,{x:targetX,y:targetY}]);
      const cleanWp = allPts.slice(1,-1);
      window.dispatchEvent(new CustomEvent("mbse:updatewaypoint",
        {detail:{id, waypoints:cleanWp}}));
    };
    const onUp=()=>{
      window.removeEventListener("mousemove",onMove);
      window.removeEventListener("mouseup",onUp);
    };
    window.addEventListener("mousemove",onMove);
    window.addEventListener("mouseup",onUp);
  },[id, segs, storedWp, pts, sourceX, sourceY, targetX, targetY]);

  return (
    <g>
      <defs>
        {/* 끝점 화살표 — 라인 길이에 비례한 크기 */}
        <marker id={mkId}
          markerWidth={arrowScale} markerHeight={arrowScale*0.6}
          refX={arrowScale-0.5} refY={arrowScale*0.3}
          orient="auto" markerUnits="strokeWidth">
          <polygon
            points={`0,0 ${arrowScale},${arrowScale*0.3} 0,${arrowScale*0.6}`}
            fill={stroke} stroke="none"/>
        </marker>
        {/* 시작점 화살표 (양방향) */}
        {ls.bidir && (
          <marker id={`${mkId}_start`}
            markerWidth={arrowScale} markerHeight={arrowScale*0.6}
            refX={0.5} refY={arrowScale*0.3}
            orient="auto-start-reverse" markerUnits="strokeWidth">
            <polygon
              points={`0,0 ${arrowScale},${arrowScale*0.3} 0,${arrowScale*0.6}`}
              fill={stroke} stroke="none"/>
          </marker>
        )}
      </defs>

      {/* 히트 영역 */}
      <path d={path} fill="none" stroke="transparent" strokeWidth={16}
        style={{cursor:"pointer"}}/>

      {/* 실제 라인 */}
      <path d={path} fill="none"
        stroke={stroke} strokeWidth={sw}
        strokeDasharray={ls.dash==="none"?"":ls.dash}
        markerEnd={`url(#${mkId})`}
        markerStart={ls.bidir ? `url(#${mkId}_start)` : undefined}
        style={{pointerEvents:"none"}}/>

      {/* 선택 시: 세그먼트 중앙 핸들 (파란 사각형) */}
      {selected && segs.map((seg, i) => {
        // 너무 짧은 세그먼트는 핸들 생략
        if (seg.len < 20) return null;
        // stub 세그먼트(첫·마지막)는 작게
        const isStub = i===0||i===segs.length-1;
        if (isStub) return null;
        const cursor = seg.isHoriz ? "ns-resize" : "ew-resize";
        return (
          <rect key={i}
            x={seg.mx-5} y={seg.my-5} width={10} height={10}
            rx={1}
            fill={selected?"#2563EB":"#3b82f6"}
            stroke="white" strokeWidth={1.5}
            style={{cursor, pointerEvents:"all"}}
            onMouseDown={e=>onSegDrag(e,i)}/>
        );
      })}

      {/* 라벨 + IC 배지 */}
      {(showLabel||icNo) && longest.len > 30 && (
        <EdgeLabelRenderer>
          <div style={{
            position:"absolute",
            transform:`translate(-50%,-100%) translate(${longest.mx}px,${longest.my}px)`,
            display:"flex",flexDirection:"column",alignItems:"center",gap:2,
            pointerEvents:"none",
          }}>
            {showLabel && (
              <div style={{
                fontSize:10, fontWeight:isSpecial?700:500,
                fontFamily:"monospace",
                background:"rgba(255,255,255,0.95)",
                padding:"1px 6px", borderRadius:4,
                border:`1px solid ${baseColor}`,
                color:baseColor,
                whiteSpace:"nowrap",
                boxShadow:"0 1px 3px rgba(0,0,0,0.1)",
              }}>{showLabel}</div>
            )}
            {icNo && (
              <div style={{
                fontSize:9,fontWeight:700,
                background:icStatus?icColor:"#64748B",
                color:"#fff",padding:"0 5px",
                borderRadius:3,whiteSpace:"nowrap",
              }}>{icNo}{icStatus?` · ${icStatus}`:""}</div>
            )}
          </div>
        </EdgeLabelRenderer>
      )}
    </g>
  );
};

// ─────────────────────────────────────────────────────────────
// SMART GUIDE — 노드 드래그 시 정렬 가상선 + 자동 스냅
// PPT처럼 다른 노드와 left/center/right/top/middle/bottom 정렬
// ─────────────────────────────────────────────────────────────
const GUIDE_TOL  = 8;   // px — 이 거리 이내면 스냅 발동
const GUIDE_COLOR = "#e11d48"; // 가상선 색상 (빨강)

// 노드의 정렬 기준점 추출 (6개: left,centerX,right,top,middleY,bottom)
const getNodeAnchors = (node) => {
  const x = node.position?.x ?? 0;
  const y = node.position?.y ?? 0;
  const w = node.width  ?? node.style?.width  ?? 120;
  const h = node.height ?? node.style?.height ?? 60;
  return {
    left:    x,
    centerX: x + w / 2,
    right:   x + w,
    top:     y,
    middleY: y + h / 2,
    bottom:  y + h,
    w, h,
  };
};

const SmartGuide = memo(({ guides }) => {
  if (!guides || guides.length === 0) return null;
  return (
    <svg style={{
      position:"absolute", inset:0,
      width:"100%", height:"100%",
      pointerEvents:"none", zIndex:9998,
      overflow:"visible",
    }}>
      {guides.map((g, i) => (
        <line key={i}
          x1={g.x1} y1={g.y1} x2={g.x2} y2={g.y2}
          stroke={GUIDE_COLOR}
          strokeWidth={1}
          strokeDasharray="4 3"
          opacity={0.85}/>
      ))}
    </svg>
  );
});
// ─────────────────────────────────────────────────────────────
const nodeTypes = { area:AreaNode, equipment:EquipmentNode, instrument:InstrumentNode, brench:BrenchNode };
const edgeTypes = { pipe:PipeEdge };

// ─────────────────────────────────────────────────────────────
// SIDEBAR
// ─────────────────────────────────────────────────────────────
const Sidebar = memo(({ onDragStart }) => {
  const [open,setOpen]=useState({ Area:true,EquipUT:true,EquipME:false,Connection:false,Instrument:false });
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
    return (
      <span style={{ display:"inline-flex",alignItems:"center",gap:1,flexShrink:0,marginRight:4 }}>
        {ls.bidir && <span style={{ color:ls.color,fontSize:9,lineHeight:1 }}>◀</span>}
        <span style={{ display:"inline-block",width:24,height:Math.min(ls.sw,3),background:ls.color,borderRadius:1 }}/>
        <span style={{ color:ls.color,fontSize:9,lineHeight:1 }}>▶</span>
      </span>
    );
  };

  return (
    <div style={{ width:195,background:"#f8fafc",borderRight:"1px solid #e2e8f0",overflowY:"auto",flexShrink:0 }}>
      <div style={{ padding:"10px 14px 7px",fontWeight:800,fontSize:13,color:"#0f172a",borderBottom:"1px solid #e2e8f0",letterSpacing:0.5 }}>Catalog</div>
      <Sec title="Area" cat="Area">
        {AREA_TYPES.map(at=>{ const c=AREA_COLORS[at]||{ bg:"rgba(219,234,254,0.45)", border:"#93c5fd", label:"#1d4ed8" }; return (
          <div key={at} draggable onDragStart={e=>onDragStart(e,"area",at)} style={{ ...iS(c?.label||"#1d4ed8"),background:c?.bg||"transparent" }} {...hov}>▭ {at}</div>
        );})}
      </Sec>
      <Sec title="Equipment (UT)" cat="EquipUT">
        {EQUIPMENT_UT.map(eq=>(
          <div key={eq} draggable onDragStart={e=>onDragStart(e,"equipment",eq)} style={iS()} {...hov}><EquipSVG type={eq} size={14}/>{eq}</div>
        ))}
      </Sec>
      <Sec title="Equipment (ME)" cat="EquipME">
        {EQUIPMENT_ME.map(eq=>(
          <div key={eq} draggable onDragStart={e=>onDragStart(e,"equipment",eq)} style={iS("#7c3aed")} {...hov}><EquipSVG type={eq} size={14}/>{eq}</div>
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
                <div style={{ display:"flex",flexDirection:"column",gap:3,marginBottom:4 }}>
                  {(d.handles||["top","bottom","left","right"]).map((dir,i)=>(
                    <div key={i} style={{ display:"flex",alignItems:"center",justifyContent:"space-between",background:"#f8fafc",border:"1px solid #e2e8f0",borderRadius:4,padding:"2px 8px" }}>
                      <span style={{ fontSize:11,color:"#334155" }}>Port {i+1}: {dir}</span>
                      <button onClick={()=>{ const h=(d.handles||["top","bottom","left","right"]).filter((_,idx)=>idx!==i); onUpdateNode(sel.id,{...d,handles:h}); }}
                        style={{ background:"none",border:"none",color:"#dc2626",cursor:"pointer",fontSize:13,padding:"0 2px" }}>×</button>
                    </div>
                  ))}
                </div>
              </>
            )}
            {/* EQUIPMENT */}
            {isNode&&sel.type==="equipment"&&(
              <>
                <label style={L}>Item No (또는 더블클릭)</label>
                <input style={I} value={d.itemNo||""} onChange={e=>upN("itemNo",e.target.value)} placeholder="e.g. RE.82.91P01"/>
                <label style={L}>설비명 (Description)</label>
                <input style={I} value={d.label||""} onChange={e=>upN("label",e.target.value)}/>

                {/* ── 설비 유형별 맞춤 사양 필드 ── */}
                <div style={{ fontWeight:600,fontSize:11,color:"#1d4ed8",margin:"8px 0 5px",borderTop:"2px solid #eff6ff",paddingTop:6,display:"flex",alignItems:"center",gap:4 }}>
                  <span>⚙ Engineering Spec</span>
                  <span style={{ fontSize:9,background:"#eff6ff",color:"#3b82f6",borderRadius:3,padding:"1px 5px",fontWeight:400 }}>
                    {d.equipType}
                  </span>
                </div>
                {getSpecFields(d.equipType).map(({key,label,unit})=>(
                  <div key={key} style={{ marginBottom:5 }}>
                    {/* 라벨: 항목명 + 단위 */}
                    <label style={{ ...L, marginBottom:1 }}>
                      {label}
                      {unit && (
                        <span style={{ color:"#94a3b8",fontWeight:400,marginLeft:4,fontSize:10 }}>
                          {unit}
                        </span>
                      )}
                    </label>
                    {/* 입력란: 단위 suffix 별도 표시, placeholder는 빈 값 */}
                    <div style={{ display:"flex",alignItems:"center",gap:4 }}>
                      <input
                        style={{ ...I, marginBottom:0, flex:1 }}
                        value={
                          // 이전 버전 잔재 키 정리: 숫자+단위 형태면 숫자만 추출하지 않고 그대로 표시
                          // 단, 값 끝에 단위가 붙어있으면 제거 후 표시
                          (() => {
                            const v = d[key] || "";
                            if (!unit || !v) return v;
                            // "80 m³/h" → "80" 으로 정리 (단위가 이미 라벨에 있으므로)
                            const stripped = v.replace(new RegExp(`\\s*${unit.replace(/[.*+?^${}()|[\]\\]/g,'\\$&')}\\s*$`), "").trim();
                            return stripped;
                          })()
                        }
                        onChange={e => upN(key, e.target.value)}
                        placeholder=""
                      />
                      {unit && (
                        <span style={{ fontSize:10,color:"#94a3b8",whiteSpace:"nowrap",flexShrink:0 }}>
                          {unit}
                        </span>
                      )}
                    </div>
                  </div>
                ))}

                {/* 포트 관리 */}
                <div style={{ fontWeight:600,fontSize:11,color:"#334155",margin:"8px 0 5px",borderTop:"1px solid #f1f5f9",paddingTop:5 }}>Port Management</div>
                <div style={{ display:"flex",flexWrap:"wrap",gap:5,marginBottom:6 }}>
                  {["top","bottom","left","right"].map(dir=>(
                    <button key={dir} onClick={()=>onAddHandle(sel.id,dir)} style={{ background:"#eff6ff",border:"1px solid #bfdbfe",color:"#1d4ed8",borderRadius:4,padding:"2px 8px",cursor:"pointer",fontSize:11 }}>+ {dir}</button>
                  ))}
                </div>
                <div style={{ display:"flex",flexDirection:"column",gap:3,marginBottom:4 }}>
                  {(d.handles||[]).map((dir,i)=>(
                    <div key={i} style={{ display:"flex",alignItems:"center",justifyContent:"space-between",background:"#f8fafc",border:"1px solid #e2e8f0",borderRadius:4,padding:"2px 8px" }}>
                      <span style={{ fontSize:11,color:"#334155" }}>Port {i+1}: {dir}</span>
                      <button onClick={()=>{ const h=(d.handles||[]).filter((_,idx)=>idx!==i); onUpdateNode(sel.id,{...d,handles:h}); }}
                        style={{ background:"none",border:"none",color:"#dc2626",cursor:"pointer",fontSize:13,padding:"0 2px" }}>×</button>
                    </div>
                  ))}
                </div>
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

                {/* 경유점(Route) 관리 */}
                {(d.waypoints?.length>0) && (
                  <div style={{ display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:7,padding:"4px 8px",background:"#fffbeb",border:"1px solid #fde68a",borderRadius:5 }}>
                    <span style={{ fontSize:11,color:"#92400e" }}>경유점 {d.waypoints.length}개</span>
                    <button onClick={()=>upE("waypoints",[])}
                      style={{ background:"none",border:"none",color:"#dc2626",cursor:"pointer",fontSize:11,fontWeight:600 }}>
                      경로 초기화
                    </button>
                  </div>
                )}
                {!(d.waypoints?.length>0) && (
                  <div style={{ fontSize:10,color:"#94a3b8",marginBottom:7 }}>
                    라인 선택 후 클릭 → 경유점 추가
                  </div>
                )}

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

                {/* IC Register 연동 정보 (IC Register Import 후 표시) */}
                {(d.ic_no||d.ic_status) && (
                  <>
                    <div style={{ fontWeight:600,fontSize:11,color:"#334155",margin:"8px 0 5px",borderTop:"1px solid #f1f5f9",paddingTop:5 }}>IC Register 정보</div>
                    <div style={{ background:"#f8fafc",border:"1px solid #e2e8f0",borderRadius:6,padding:"8px 10px",fontSize:11 }}>
                      {[
                        ["IC No.",    d.ic_no],
                        ["ICD 번호",  d.icd_no],
                        ["상태",      d.ic_status],
                        ["우선순위",  d.ic_priority],
                        ["마감일",    d.ic_due],
                        ["완료일",    d.ic_closed],
                        ["담당 (From)", d.ic_resp_from],
                        ["담당 (To)",   d.ic_resp_to],
                        ["비고",        d.ic_remark],
                      ].filter(([,v])=>v).map(([k,v])=>(
                        <div key={k} style={{ display:"flex",justifyContent:"space-between",marginBottom:3 }}>
                          <span style={{ color:"#64748b",fontWeight:500 }}>{k}</span>
                          <span style={{
                            fontWeight:700,
                            color: k==="상태"
                              ? (v==="CLOSED"?"#16a34a":v==="OVERDUE"?"#dc2626":v==="IN PROGRESS"?"#2563eb":"#ca8a04")
                              : "#1e293b"
                          }}>{v}</span>
                        </div>
                      ))}
                    </div>
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
// ─────────────────────────────────────────────────────────────
// Item No. 정규화 — 매칭 효율 개선
// 규칙:
//  1) 앞의 "RE" 제거
//  2) 마침표(.) 제거
//  3) A/B/C suffix (끝에 오는 단독 대문자) 제거 후 별도 보관
//  4) 숫자+대문자알파벳+숫자 패턴만 남겨서 비교
//
// 예) "RE.82.91P01 A/B" → "8291P01"
//     "82.91P01"         → "8291P01"
//     "RE.81.1P01"       → "811P01"
//     "82.93H01 A/B"     → "8293H01"
// ─────────────────────────────────────────────────────────────
const normalizeItemNo = (raw) => {
  if (!raw) return "";
  let s = raw.toUpperCase().trim();
  // 1) 선두 "RE" 제거 (RE. 또는 RE 단독)
  s = s.replace(/^RE\.?/, "");
  // 2) 공백 이후 단독 대문자 suffix 제거: " A/B", " A", " B", " C" 등
  s = s.replace(/\s+[A-Z](\/[A-Z])*\s*$/, "");
  // 3) 마침표, 하이픈, 공백 모두 제거
  s = s.replace(/[\.\-\s]/g, "");
  // 4) 숫자+대문자+숫자 핵심 패턴 추출
  //    예: "8291P01" (숫자들 + 대문자 1개 이상 + 숫자들)
  const m = s.match(/(\d+[A-Z]+\d*|\d+[A-Z]+)/);
  return m ? m[0] : s;
};

// Item No. 배열 정규화 (A/B/C 분리 포함)
const normalizeItemNos = (rawItemNo) => {
  const base = rawItemNo.trim();
  const results = new Set();

  // 원본 정규화
  results.add(normalizeItemNo(base));

  // A/B/C/D suffix 분리: "RE.82.91P01 A/B" → A, B 각각
  const abcMatch = base.match(/^(.+?)\s+([A-Z](?:\/[A-Z])+)$/);
  if (abcMatch) {
    abcMatch[2].split("/").forEach(s => {
      results.add(normalizeItemNo(`${abcMatch[1]} ${s}`));
    });
  }
  // 숫자만으로 이루어진 suffix 처리: "RE.82.91P01/06" 형태
  const numSuffix = base.match(/^(.+?)\/(\d+)$/);
  if (numSuffix) {
    results.add(normalizeItemNo(numSuffix[1]));
  }

  return [...results].filter(Boolean);
};

// ─────────────────────────────────────────────────────────────
// SPEC IMPORT MODAL
// 사양서 텍스트 붙여넣기 → Item No. 기준 자동 매핑
// ─────────────────────────────────────────────────────────────
const SpecImportModal = ({ nodes, onApply, onCancel }) => {
  const [text,   setText]   = useState("");
  const [parsed, setParsed] = useState(null);
  const [status, setStatus] = useState("idle"); // idle|parsing|done|error

  // ── 사양서 파싱 ────────────────────────────────────────────
  // Item No.가 "RE.xx.xxXxx" 형태이거나 "Item No. : xxx" 패턴으로 등장
  const parseSpec = (raw) => {
    const blocks = [];
    // 패턴: "Item No." 또는 "Item No :" 로 구분하여 블록 분리
    const parts = raw.split(/(?=Item\s*No\.?\s*[\:：]?\s*[A-Za-z0-9])/i)
      .filter(p => /Item\s*No/i.test(p));

    parts.forEach(block => {
      // Item No. 추출
      const itemMatch = block.match(/Item\s*No\.?\s*[\:：]?\s*([A-Za-z0-9\.\-\/\s]+?)(?:\n|\r|설\s*비\s*명)/i);
      if (!itemMatch) return;
      const rawItemNo = itemMatch[1].trim().replace(/\s+/g," ");
      // A/B/C suffix 처리 + 정규화
      const itemNos = [rawItemNo];
      const abcMatch = rawItemNo.match(/^(.+?)\s+([A-Z](?:\/[A-Z])+)$/);
      if (abcMatch) {
        abcMatch[2].split("/").forEach(s => itemNos.push(`${abcMatch[1]} ${s}`));
      }
      // 정규화된 키 목록 (매칭용)
      const normalizedKeys = normalizeItemNos(rawItemNo);

      // 설비명
      const nameMatch = block.match(/설\s*비\s*명\s*[\:：]?\s*(.+?)(?:\n|\r)/);
      const equipName = nameMatch?.[1]?.trim() || "";

      // 수량
      const qtyMatch = block.match(/수\s*량\s*[\:：]?\s*(\d+)/);

      // 사양 추출 함수
      const extractVal = (patterns) => {
        for (const p of patterns) {
          const m = block.match(new RegExp(p + "\\s*[\\:：]?\\s*([^\\n\\r]+)","i"));
          if (m) {
            const v = m[1].trim().replace(/\s*\/\s*/g," / ");
            if (v && v !== "-" && v !== "TBD") return v;
          }
        }
        return "";
      };

      const spec = {
        itemNos,
        normalizedKeys,   // ← 매칭용 정규화 키
        equipName,
        quantity: qtyMatch?.[1] || "1",
        // 공통
        kindOfLiquid:   extractVal(["Kind of [Ll]iquid","Flow medium","Medium"]),
        capacity:       extractVal(["Capacity","Volume flow rate","Total capacity","Flow rate from process","Flow rate\\s*\\(m"]),
        deliveryHead:   extractVal(["Delivery head"]),
        suctionPress:   extractVal(["Suction [Pp]ress"]),
        dischargePress: extractVal(["Discharge [Pp]ress"]),
        liquidTemp:     extractVal(["Liquid [Tt]emp","Inlet [Tt]emp"]),
        designP:        extractVal(["Design [Pp]ress"]),
        designT:        extractVal(["Design [Tt]emp","Design temp"]),
        operPress:      extractVal(["Operating [Pp]ress","Operating pressure"]),
        operTemp:       extractVal(["Operating [Tt]emp","Operating temperature"]),
        construction:   extractVal(["Type of construction","Type of design"]),
        shaftSeal:      extractVal(["Type of shaft seal","Shaft seal"]),
        material:       extractVal(["Casing\t\t","Casing\\s+:", "Shell material","Framework"]),
        impeller:       extractVal(["Impeller\t\t","Impeller\\s+:"]),
        // HX 전용
        flowRatePri:    extractVal(["Flow rate \\(m.*\\).*Primary","Primary.*flow"]),
        inletTempPri:   extractVal(["Inlet.*temp.*Primary","Primary.*inlet"]),
        outletTempPri:  extractVal(["Outlet.*temp.*Primary"]),
        inletTempSec:   extractVal(["Inlet.*temp.*Secondary","Secondary.*inlet"]),
        outletTempSec:  extractVal(["Outlet.*temp.*Secondary"]),
        // Cooling Tower
        outletTemp:     extractVal(["Outlet [Tt]emp"]),
        wetBulbTemp:    extractVal(["Wet bulb [Tt]emp"]),
        numCells:       extractVal(["Number of cells"]),
        // Filter
        meshSize:       extractVal(["Mesh size"]),
        pressureDrop:   extractVal(["Press.*drop"]),
        // Clarifier
        overflowSolid:  extractVal(["Overflow [Ss]olid"]),
        underflowSolid: extractVal(["Underflow [Ss]olid"]),
        // Chemical dosing
        tankVolume:     extractVal(["Tank volume"]),
        pumpType:       extractVal(["Type of pump"]),
        // Compressor / Fan
        inletPress:     extractVal(["Inlet [Pp]ress"]),
        outletTemp2:    extractVal(["Outlet [Tt]emp.*max"]),
        numStages:      extractVal(["Number of stages"]),
        staticPress:    extractVal(["Static [Pp]ress"]),
        noiseLevel:     extractVal(["Noise level"]),
        // Misc
        medium:         extractVal(["Medium\t\t","Medium\\s+:"]),
        installation:   extractVal(["Place of installation"]),
        solidContent:   extractVal(["Solid content","Density of solid"]),
        particleSize:   extractVal(["Particle size"]),
      };
      // 빈 값 제거
      Object.keys(spec).forEach(k => {
        if (typeof spec[k] === "string" && !spec[k]) delete spec[k];
      });
      blocks.push(spec);
    });
    return blocks;
  };

  const handleParse = () => {
    if (!text.trim()) return;
    setStatus("parsing");
    setTimeout(() => {
      try {
        const blocks = parseSpec(text);
        setParsed(blocks);
        setStatus("done");
      } catch(e) {
        setStatus("error");
      }
    }, 100);
  };

  // 노드 Item No.와 매칭
  const getMatchedNodes = (block) => {
    const blockKeys = block.normalizedKeys || normalizeItemNos(block.itemNos?.[0]||"");
    return nodes.filter(n => {
      if (n.type !== "equipment") return false;
      const nodeKey = normalizeItemNo(n.data?.itemNo||"");
      if (!nodeKey) return false;
      return blockKeys.some(bk => bk && bk === nodeKey);
    });
  };

  const totalMatches = parsed ? parsed.reduce((s,b)=>s+getMatchedNodes(b).length,0) : 0;

  return (
    <div style={{ position:"fixed",inset:0,background:"rgba(0,0,0,0.55)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:9999 }}>
      <div style={{ background:"#fff",borderRadius:12,padding:0,width:640,maxHeight:"85vh",display:"flex",flexDirection:"column",boxShadow:"0 20px 60px rgba(0,0,0,0.25)" }}>
        {/* 헤더 */}
        <div style={{ padding:"16px 20px 12px",borderBottom:"1px solid #e2e8f0",display:"flex",justifyContent:"space-between",alignItems:"center" }}>
          <div>
            <div style={{ fontWeight:800,fontSize:15,color:"#0f172a" }}>📋 사양서 자동 입력</div>
            <div style={{ fontSize:11,color:"#64748b",marginTop:2 }}>사양서 텍스트를 붙여넣으면 Item No. 기준으로 자동 매핑됩니다</div>
          </div>
          <button onClick={onCancel} style={{ background:"#f1f5f9",border:"none",borderRadius:5,padding:"4px 10px",cursor:"pointer",color:"#64748b" }}>✕ 닫기</button>
        </div>

        <div style={{ flex:1,overflowY:"auto",padding:"16px 20px" }}>
          {status !== "done" ? (
            <>
              <label style={{ fontSize:11,fontWeight:600,color:"#334155",display:"block",marginBottom:6 }}>
                사양서 텍스트 붙여넣기 (Ctrl+V)
              </label>
              <textarea
                style={{ width:"100%",height:220,border:"1.5px solid #e2e8f0",borderRadius:8,padding:"10px",fontSize:11,resize:"vertical",boxSizing:"border-box",fontFamily:"monospace",lineHeight:1.5 }}
                value={text}
                onChange={e=>setText(e.target.value)}
                placeholder={"Item No.\tRE.81.1P01 A/B\n\n설 비 명\tMake up water pump\n수    량\t2 sets\n\n설비사양\n\t설계조건\n\tCapacity (m3/h)\t: 80\n\tDelivery head (m)\t: 140\n\t...\n\n여러 설비를 한번에 붙여넣기 가능합니다."}
              />
              {status === "error" && (
                <div style={{ color:"#dc2626",fontSize:11,marginTop:6 }}>파싱 오류가 발생했습니다. 텍스트 형식을 확인해주세요.</div>
              )}
            </>
          ) : (
            <>
              <div style={{ display:"flex",alignItems:"center",gap:8,marginBottom:12,padding:"8px 12px",background:"#f0fdf4",borderRadius:6,border:"1px solid #bbf7d0" }}>
                <span style={{ fontSize:16 }}>✅</span>
                <div>
                  <div style={{ fontWeight:700,fontSize:12,color:"#15803d" }}>
                    {parsed.length}개 설비 파싱 완료 — MBSE 모델 매칭: {totalMatches}개 노드
                  </div>
                  <div style={{ fontSize:10,color:"#16a34a" }}>매칭된 노드에 사양이 자동 입력됩니다</div>
                </div>
              </div>

              {parsed.map((block, bi) => {
                const matched = getMatchedNodes(block);
                return (
                  <div key={bi} style={{ marginBottom:10,border:`1px solid ${matched.length?"#bfdbfe":"#e2e8f0"}`,borderRadius:8,overflow:"hidden" }}>
                    <div style={{ padding:"8px 12px",background:matched.length?"#eff6ff":"#f8fafc",display:"flex",justifyContent:"space-between",alignItems:"center" }}>
                      <div>
                        <span style={{ fontWeight:700,fontSize:12,color:"#1d4ed8" }}>{block.itemNos[0]}</span>
                        <span style={{ fontSize:11,color:"#64748b",marginLeft:8 }}>{block.equipName}</span>
                      </div>
                      {matched.length > 0
                        ? <span style={{ fontSize:10,background:"#1d4ed8",color:"#fff",borderRadius:4,padding:"1px 7px",fontWeight:600 }}>✓ {matched.length}개 매칭</span>
                        : <span style={{ fontSize:10,background:"#f1f5f9",color:"#94a3b8",borderRadius:4,padding:"1px 7px" }}>미매칭</span>
                      }
                    </div>
                    <div style={{ padding:"8px 12px",display:"flex",flexWrap:"wrap",gap:"4px 12px" }}>
                      {Object.entries(block).filter(([k])=>!["itemNos","equipName","quantity"].includes(k)).map(([k,v])=>(
                        <div key={k} style={{ fontSize:10,color:"#475569" }}>
                          <span style={{ color:"#94a3b8" }}>{k}:</span> <span style={{ fontWeight:600 }}>{v}</span>
                        </div>
                      ))}
                    </div>
                  </div>
                );
              })}
            </>
          )}
        </div>

        {/* 하단 버튼 */}
        <div style={{ padding:"12px 20px",borderTop:"1px solid #e2e8f0",display:"flex",gap:8,justifyContent:"flex-end" }}>
          {status !== "done" ? (
            <>
              <button onClick={onCancel} style={{ background:"#f1f5f9",border:"1px solid #e2e8f0",borderRadius:6,padding:"8px 16px",cursor:"pointer",fontSize:12 }}>취소</button>
              <button onClick={handleParse}
                disabled={!text.trim()}
                style={{ background:text.trim()?"#1d4ed8":"#94a3b8",color:"#fff",border:"none",borderRadius:6,padding:"8px 20px",cursor:text.trim()?"pointer":"default",fontWeight:700,fontSize:12 }}>
                {status==="parsing"?"파싱 중...":"🔍 분석 시작"}
              </button>
            </>
          ) : (
            <>
              <button onClick={()=>{setParsed(null);setStatus("idle");setText("");}} style={{ background:"#f1f5f9",border:"1px solid #e2e8f0",borderRadius:6,padding:"8px 16px",cursor:"pointer",fontSize:12 }}>다시 붙여넣기</button>
              <button onClick={()=>onApply(parsed)} disabled={totalMatches===0}
                style={{ background:totalMatches?"#1d4ed8":"#94a3b8",color:"#fff",border:"none",borderRadius:6,padding:"8px 20px",cursor:totalMatches?"pointer":"default",fontWeight:700,fontSize:12 }}>
                ✅ {totalMatches}개 노드에 사양 적용
              </button>
            </>
          )}
        </div>
      </div>
    </div>
  );
};

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
  const [guides,setGuides] = useState([]);
  const [showSpecImport, setShowSpecImport] = useState(false); // Smart Guide 가상선
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

  // ── 인라인 편집 + waypoint 이벤트 수신 ───────────────────
  useEffect(()=>{
    const fn=e=>{
      const {id,label,itemNo,toggleIO,waypoints}=e.detail;
      setNodes(ns=>ns.map(n=>{
        if(n.id!==id) return n;
        if(label!==undefined)    return { ...n,data:{ ...n.data,label } };
        if(itemNo!==undefined)   return { ...n,data:{ ...n.data,itemNo } };
        if(toggleIO!==undefined) return { ...n,data:{ ...n.data,showIO:toggleIO } };
        return n;
      }));
      // waypoint → edge 업데이트 (정규화 적용)
      if(waypoints!==undefined){
        setEdges(es=>es.map(e=>{
          if(e.id!==id) return e;
          // 전체 pts 기준으로 정규화
          return { ...e, data:{ ...e.data, waypoints } };
        }));
      }
    };
    window.addEventListener("mbse:updatelabel",fn);
    window.addEventListener("mbse:updatewaypoint",fn);
    return()=>{
      window.removeEventListener("mbse:updatelabel",fn);
      window.removeEventListener("mbse:updatewaypoint",fn);
    };
  },[setNodes,setEdges]);

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
      if(sub==="Brench"){
        // ── Brench를 배관 위에 드롭하면 자동 분기 ──────────
        // 엣지의 실제 경로 midpoint로 근접 계산 (노드 좌표 기반보다 정확)
        const SNAP = 60; // 픽셀 반경
        let nearEdge = null;
        let minDist   = SNAP;

        edges.forEach(edge => {
          const srcN = nodes.find(n=>n.id===edge.source);
          const tgtN = nodes.find(n=>n.id===edge.target);
          if(!srcN||!tgtN) return;
          // 엣지의 orthogonal 경로 포인트들
          const sx=srcN.position.x+(srcN.width||110)/2;
          const sy=srcN.position.y+(srcN.height||64)/2;
          const tx=tgtN.position.x+(tgtN.width||110)/2;
          const ty=tgtN.position.y+(tgtN.height||64)/2;
          const wp=edge.data?.waypoints||[];
          const pts=[{x:sx,y:sy},...wp,{x:tx,y:ty}];
          // 각 세그먼트까지 거리
          for(let i=0;i<pts.length-1;i++){
            const a=pts[i],b=pts[i+1];
            const dx=b.x-a.x,dy=b.y-a.y,len=dx*dx+dy*dy||1;
            const t=Math.max(0,Math.min(1,((pos.x-a.x)*dx+(pos.y-a.y)*dy)/len));
            const d=Math.hypot(pos.x-(a.x+t*dx),pos.y-(a.y+t*dy));
            if(d<minDist){minDist=d;nearEdge=edge;}
          }
        });

        if(nearEdge){
          const brId=uid("br");
          const eData=nearEdge.data||{};
          // Brench 노드 생성
          setNodes(ns=>[...ns,{id:brId,type:"brench",position:pos,data:{}}]);
          // 기존 엣지 제거 → source→brench / brench→target 2개 생성
          setEdges(es=>[
            ...es.filter(e=>e.id!==nearEdge.id),
            {id:uid("e"),type:"pipe",
             source:nearEdge.source,target:brId,
             data:{...eData,waypoints:[]}},
            {id:uid("e"),type:"pipe",
             source:brId,target:nearEdge.target,
             data:{...eData,waypoints:[]}},
          ]);
        } else {
          setNodes(ns=>[...ns,{id:uid("br"),type:"brench",position:pos,data:{}}]);
        }
      } else {
        setNodes(ns=>[...ns,{id:uid("br"),type:"brench",position:pos,data:{_hint:sub}}]);
      }
    }
  },[screenToFlowPosition,setNodes,setEdges,edges,nodes]);

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

  // ── Smart Guide: 드래그 중 정렬선 계산 + 스냅 ──────────────
  const onNodeDrag = useCallback((_evt, dragNode) => {
    const others = nodes.filter(n => n.id !== dragNode.id && n.type !== "area");
    if (others.length === 0) { setGuides([]); return; }

    const da = getNodeAnchors(dragNode);
    const newGuides = [];
    let snapX = null, snapY = null;

    others.forEach(other => {
      const oa = getNodeAnchors(other);
      const canvasW = 16000, canvasH = 16000;

      // ── 수직 정렬 (x 기준) ──────────────────────────────
      const xPairs = [
        [da.left,    oa.left,    "left-left"],
        [da.left,    oa.centerX, "left-center"],
        [da.left,    oa.right,   "left-right"],
        [da.centerX, oa.left,    "center-left"],
        [da.centerX, oa.centerX, "center-center"],
        [da.centerX, oa.right,   "center-right"],
        [da.right,   oa.left,    "right-left"],
        [da.right,   oa.centerX, "right-center"],
        [da.right,   oa.right,   "right-right"],
      ];
      xPairs.forEach(([dv, ov]) => {
        if (Math.abs(dv - ov) <= GUIDE_TOL) {
          if (snapX === null) snapX = ov - (dv - da.left) ; // 보정값
          newGuides.push({
            x1: ov, y1: 0, x2: ov, y2: canvasH,
          });
        }
      });

      // ── 수평 정렬 (y 기준) ──────────────────────────────
      const yPairs = [
        [da.top,     oa.top,     "top-top"],
        [da.top,     oa.middleY, "top-middle"],
        [da.top,     oa.bottom,  "top-bottom"],
        [da.middleY, oa.top,     "middle-top"],
        [da.middleY, oa.middleY, "middle-middle"],
        [da.middleY, oa.bottom,  "middle-bottom"],
        [da.bottom,  oa.top,     "bottom-top"],
        [da.bottom,  oa.middleY, "bottom-middle"],
        [da.bottom,  oa.bottom,  "bottom-bottom"],
      ];
      yPairs.forEach(([dv, ov]) => {
        if (Math.abs(dv - ov) <= GUIDE_TOL) {
          if (snapY === null) snapY = ov - (dv - da.top);
          newGuides.push({
            x1: 0, y1: ov, x2: canvasW, y2: ov,
          });
        }
      });
    });

    setGuides(newGuides);

    // 스냅 적용
    if (snapX !== null || snapY !== null) {
      setNodes(ns => ns.map(n => {
        if (n.id !== dragNode.id) return n;
        return {
          ...n,
          position: {
            x: snapX !== null ? snapX : n.position.x,
            y: snapY !== null ? snapY : n.position.y,
          }
        };
      }));
    }
  }, [nodes, setNodes]);

  // ── 사양서 자동 입력 적용 ─────────────────────────────────
  const applySpecImport = useCallback((parsedBlocks) => {
    setNodes(ns => ns.map(n => {
      if (n.type !== "equipment") return n;
      const nodeKey = normalizeItemNo(n.data?.itemNo||"");
      if (!nodeKey) return n;
      // 정규화 키로 매칭 블록 찾기
      const block = parsedBlocks.find(b => {
        const bKeys = b.normalizedKeys || normalizeItemNos(b.itemNos?.[0]||"");
        return bKeys.some(bk => bk && bk === nodeKey);
      });
      if (!block) return n;
      // 사양 필드 업데이트 (itemNos, normalizedKeys, equipName, quantity 제외)
      const specData = {};
      Object.entries(block).forEach(([k,v]) => {
        if (!["itemNos","normalizedKeys","equipName","quantity"].includes(k)) specData[k] = v;
      });
      // 설비명도 업데이트 (기존 값 없을 때만)
      if (!n.data.label && block.equipName) specData.label = block.equipName;
      return { ...n, data: { ...n.data, ...specData } };
    }));
    setShowSpecImport(false);
    setSaveMsg(`✅ 사양서 ${parsedBlocks.length}건 적용 완료`);
    setTimeout(()=>setSaveMsg(""),3000);
  },[setNodes]);

  const onNodeDragStop = useCallback((_, dragNode) => {    setGuides([]);
    // 이동된 노드에 연결된 엣지의 waypoints 초기화 → 자동 재라우팅
    setEdges(es => es.map(e => {
      if (e.source === dragNode.id || e.target === dragNode.id) {
        return { ...e, data: { ...e.data, waypoints: [] } };
      }
      return e;
    }));
  }, [setEdges]);

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
      setXlsxMsg("IC Register 생성 중...");
      const result = await exportToExcel(nodes, edges);
      setXlsxMsg(`✅ IC ${result.icCount}건 / 설비 ${result.equipCount}건 Export 완료`);
      setTimeout(()=>setXlsxMsg(""),3500);
    } catch(err) {
      setXlsxMsg("오류: " + err.message);
      setTimeout(()=>setXlsxMsg(""), 3000);
    }
  };

  // Excel Import
  const onExcelImport = async (e) => {
    const f = e.target.files[0]; if(!f) return;
    try {
      setXlsxMsg("IC Register 불러오는 중...");
      const result = await importFromExcel(f, nodes, edges, setNodes, setEdges);
      setXlsxMsg(`✅ ${result.msg}`);
      setTimeout(()=>setXlsxMsg(""), 3000);
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

          {/* 📋 사양서 자동 입력 */}
          <button onClick={()=>setShowSpecImport(true)}
            style={{ background:"#7c3aed",color:"#fff",border:"none",borderRadius:5,padding:"3px 10px",cursor:"pointer",fontSize:11,fontWeight:700 }}>
            📋 사양서 Import
          </button>
          <button onClick={()=>fitView({ padding:0.15, duration:400 })}
            style={{ background:"#1e293b",color:"#94a3b8",border:"1px solid #334155",borderRadius:5,padding:"3px 10px",cursor:"pointer",fontSize:11 }}>
            ⊞ 전체보기
          </button>

          {/* Excel / IC Register */}
          <button onClick={onExcelExport} style={{ background:"#14532d",color:"#86efac",border:"1px solid #166534",borderRadius:5,padding:"3px 10px",cursor:"pointer",fontSize:11,fontWeight:600 }}>📊 IC Register ⬇</button>
          <button onClick={()=>xlsxRef.current?.click()} style={{ background:"#14532d",color:"#86efac",border:"1px solid #166534",borderRadius:5,padding:"3px 10px",cursor:"pointer",fontSize:11,fontWeight:600 }}>📊 IC Register ⬆</button>
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
              onNodeDrag={onNodeDrag}
              onNodeDragStop={onNodeDragStop}
              nodeTypes={nodeTypes} edgeTypes={edgeTypes}
              connectionMode={ConnectionMode.Loose}
              connectionLineType="straight"
              connectionLineStyle={{ stroke:"#0d9488",strokeWidth:2,strokeDasharray:"4 2" }}
              defaultEdgeOptions={{ type:"pipe" }}
              // ── fitView: 전체 노드가 화면에 맞게 보임 ─────────
              fitView
              fitViewOptions={{ padding:0.15, includeHiddenNodes:false }}
              minZoom={0.05}
              maxZoom={2}
              snapToGrid snapGrid={[10,10]}
              deleteKeyCode={null}
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
              {/* Smart Guide 가상선 오버레이 */}
              {guides.length > 0 && (
                <Panel position="top-left" style={{margin:0,padding:0,pointerEvents:"none",width:"100%",height:"100%"}}>
                  <SmartGuide guides={guides}/>
                </Panel>
              )}
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

      {/* 사양서 자동 입력 모달 */}
      {showSpecImport && (
        <SpecImportModal
          nodes={nodes}
          onApply={applySpecImport}
          onCancel={()=>setShowSpecImport(false)}
        />
      )}

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
