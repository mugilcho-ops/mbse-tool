// ============================================================
// MBSE Interface Master v10
// EM System: Interface Type 분류 + ICD/TQ 관리 + Scope/Vendor + 9단계 IC Status
// Excel 6 Sheets: IC / ICD / Equipment / Connection / Scope / Requirements
// React + ReactFlow  |  package.json: "reactflow": "^11.11.4"
// ============================================================

import React, { useState, useCallback, useRef, useEffect, useMemo, memo } from "react";
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
  useStore,
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

  // ══════════════════════════════════════════════════════════
  // v10.1 SHEET 0: Scope of Supply (담당자 배정표)
  // Plant / Package / Equipment 계층별로
  //  - POSCO 담당 (PM·운전·정비, 복수)
  //  - Supplier 담당 (대표·PM·Mech·Piping·Process·EIC·Civil·Arch)
  // ══════════════════════════════════════════════════════════
  const sosRows = [];
  // 담당자 객체 → 문자열 ("PM: 홍길동 / 운전: 김철수")
  const staffToStr = (obj, roles) => {
    if (!obj) return "";
    return roles.filter(r=>obj[r]).map(r=>`${r}: ${obj[r]}`).join(" / ");
  };
  // sos 데이터 추출 헬퍼
  const sosEntry = (n, level) => {
    const d = n.data || {};
    const sos = d.sos || {};   // {posco:{role:names}, supplier:{role:names}}
    const row = {
      "구분":      level,
      "이름":      d.label || d.itemNo || "",
      "Type":      d.areaType || d.equipType || n.type || "",
      "Vendor":    d.vendorName || "",
    };
    // POSCO 담당 (역할별 컬럼)
    SOS_POSCO_ROLES.forEach(r => {
      row[`POSCO·${r}`] = sos.posco?.[r] || "";
    });
    // Supplier 담당 (역할별 컬럼)
    SOS_SUPPLIER_ROLES.forEach(r => {
      row[`Supplier·${r}`] = sos.supplier?.[r] || "";
    });
    row["Node ID"] = n.id;
    return row;
  };

  // Plant → Package(System) → Equipment 순서로 정렬
  const plantAreas   = areaNodes.filter(a => a.data?.areaType === "Plant");
  const systemAreas  = areaNodes.filter(a => a.data?.areaType === "System");
  const packageAreas = areaNodes.filter(a => a.data?.areaType === "Package");
  const itemAreas    = areaNodes.filter(a => a.data?.areaType === "Item");
  const equipNodes   = nodes.filter(n => n.type === "equipment");

  plantAreas.forEach(a   => sosRows.push(sosEntry(a, "Plant")));
  systemAreas.forEach(a  => sosRows.push(sosEntry(a, "System")));
  packageAreas.forEach(a => sosRows.push(sosEntry(a, "Package")));
  itemAreas.forEach(a    => sosRows.push(sosEntry(a, "Item")));
  equipNodes.forEach(n   => sosRows.push(sosEntry(n, "Equipment")));


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
        "상태":             d.ic_status || "OPEN",
        "우선순위":         fluidSub === "" ? "Medium" : "High",
        "등록일":           new Date().toISOString().slice(0,10),
        "목표 완료일":      d.ic_due    || "",
        "실제 완료일":      d.ic_closed || "",
        "담당자 (From)":    d.ic_resp_from || "",
        "담당자 (To)":      d.ic_resp_to   || "",
        "비고":             d.ic_remark || "",
        "IF Type":          d.ifType    || "",
        "ICD 번호":         d.icd_no    || `ICD-${String(icRows.length+1).padStart(3,"0")}`,
        "ICD 상태":         d.icd_status|| "",
        "TQ 번호":          d.tq_no     || "",
        "TQ 상태":          d.tq_status || "",
        "Open Items":       d.openItems || "",
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
      "상태":             d.ic_status || "OPEN",
      "우선순위":         "Medium",
      "등록일":           new Date().toISOString().slice(0,10),
      "목표 완료일":      d.ic_due    || "",
      "실제 완료일":      d.ic_closed || "",
      "담당자 (From)":    d.ic_resp_from || "",
      "담당자 (To)":      d.ic_resp_to   || "",
      "비고":             d.ic_remark || "",
      "IF Type":          d.ifType    || "",
      "ICD 번호":         d.icd_no    || `ICD-${String(icRows.length+1).padStart(3,"0")}`,
      "ICD 상태":         d.icd_status|| "",
      "TQ 번호":          d.tq_no     || "",
      "TQ 상태":          d.tq_status || "",
      "Open Items":       d.openItems || "",
      "Edge ID":          e.id,
    });
  });

  // ══════════════════════════════════════════════════════════
  // v10 SHEET: ICD Register (ICD 번호 기준 집계)
  // ══════════════════════════════════════════════════════════
  const icdRows = [];
  edges.forEach(e => {
    const d = e.data || {};
    if (!d.icd_no) return; // ICD 번호 있는 것만
    const srcNode = nodes.find(n => n.id === e.source);
    const tgtNode = nodes.find(n => n.id === e.target);
    const srcArea = srcNode ? getAreaOf(srcNode, nodes) : null;
    const tgtArea = tgtNode ? getAreaOf(tgtNode, nodes) : null;
    icdRows.push({
      "ICD 번호":      d.icd_no,
      "ICD 상태":      d.icd_status || "",
      "IF Type":       d.ifType || "",
      "Sender (Package)":   srcArea?.data?.vendorName || srcArea?.data?.label || "",
      "Receiver (Package)": tgtArea?.data?.vendorName || tgtArea?.data?.label || "",
      "IC 상태":       d.ic_status || "",
      "IFA Date":      d.icd_ifaDate || "",
      "IFC Date":      d.icd_ifcDate || "",
      "TQ 번호":       d.tq_no || "",
      "TQ 상태":       d.tq_status || "",
      "Open Items":    d.openItems || "",
      "Edge ID":       e.id,
    });
  });

  // ══════════════════════════════════════════════════════════
  // v10.3 SHEET: IT Register (Interface Tie-in — Area↔Area)
  // ══════════════════════════════════════════════════════════
  const itRows = [];
  edges.filter(e => e.type === "itEdge").forEach(e => {
    const d = e.data || {};
    const srcNode = nodes.find(n => n.id === e.source);
    const tgtNode = nodes.find(n => n.id === e.target);
    const srcName = srcNode?.data?.label || srcNode?.id || "";
    const tgtName = tgtNode?.data?.label || tgtNode?.id || "";
    // 공급 구분별 → "AreaA:공급처A / AreaB:공급처B"
    const supplyStr = (d.supplyByArea||[])
      .map(s => `${s.areaName}: ${s.supplier}`).join(" / ");
    // 영향인자 → "내용 [진행현황|담당자]" 형식으로 "; " join (구버전 문자열 호환)
    // 영향인자 → "내용 [진행현황|담당자]" 형식, 셀 내 줄바꿈(\n)으로 구분
    // (wrapText 셀이라 엑셀에서 항목별 줄로 표시 — Alt+Enter로 항목 추가/구분 가능)
    const influStr = (d.influenceFactors||[]).map(f => {
      const n = typeof f === "string" ? { text:f, status:"확인 중", owner:"" } : f;
      const meta = [n.status, n.owner].filter(Boolean).join("|");
      return meta ? `${n.text} [${meta}]` : n.text;
    }).join("\n");
    // 체크리스트 카테고리별 펼치기
    const cl = d.checklist || {};
    const row = {
      "IT No.":                 d.itNo || "",
      "IC&TI Description":      `${srcName} ↔ ${tgtName}`,
      "Source Area":            srcName,
      "Target Area":            tgtName,
      "공급 구분별 Inter-connection": supplyStr,
      "기능 운영간 Inter-connection": d.functionalDesc || "",
      "Process 간 주요 영향인자":  influStr,
    };
    IT_CHECKLIST_CATEGORIES.forEach(cat => {
      const v = normalizeITChecklistCat(cl[cat.key]);
      row[`${cat.label} (공급)`]     = v.supplier || "";
      row[`${cat.label} (공급담당)`] = v.supplierOwner || "";
      row[`${cat.label} (주관)`]     = v.responsible || "";
      row[`${cat.label} (주관담당)`] = v.responsibleOwner || "";
      // 항목 → "내용 [진행현황]" 줄바꿈 join (wrapText 셀, Alt+Enter로 항목 추가 가능)
      row[`${cat.label} (내용)`]     = (v.items||[])
        .filter(it => it.text.trim())
        .map(it => it.status && it.status !== "확인 중" ? `${it.text} [${it.status}]` : `${it.text} [${it.status||"확인 중"}]`)
        .join("\n");
    });
    row["Edge ID"] = e.id;
    itRows.push(row);
  });

  // ══════════════════════════════════════════════════════════
  // v10 SHEET: Scope Register (Area 노드 Scope/Vendor 정보)
  // ══════════════════════════════════════════════════════════
  const scopeRows = [];
  nodes.filter(n => n.type === "area").forEach(n => {
    const d = n.data || {};
    if (!d.wbsCode && !d.vendorName && !d.scopeText) return; // 입력된 것만
    // POSCO/Eng 담당자 합치기
    const poscoStr = d.poscoStaff
      ? Object.entries(d.poscoStaff).filter(([,v])=>v).map(([k,v])=>`${k}: ${v}`).join(" / ")
      : "";
    const engStr = d.engStaff
      ? Object.entries(d.engStaff).filter(([,v])=>v).map(([k,v])=>`${k}: ${v}`).join(" / ")
      : "";
    scopeRows.push({
      "WBS Code":       d.wbsCode || "",
      "Area":           d.label || "",
      "Type":           d.areaType || "",
      "Team":           d.teamId || "",
      "Discipline":     d.discipline || "",
      "범위 기술":       d.scopeText || "",
      "포함 범위":       d.inclusions || "",
      "제외 범위":       d.exclusions || "",
      "Vendor 회사":     d.vendorName || "",
      "Vendor 국가":     d.vendorCountry || "",
      "계약번호":        d.vendorContract || "",
      "Interface Coordinator": d.vendorICA || "",
      "POSCO 담당":      poscoStr,
      "Engineering 담당": engStr,
      "Node ID":        n.id,
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
    // 소속 Area 이름만 (title 태그 제외)
    const areaLabel = area?.data?.label || "";
    (d.requirements || []).forEach(r => {
      reqRows.push({
        "Node ID":      n.id,
        "Item No.":     d.itemNo || d.label || "",
        "소속 Area":     areaLabel,
        "Stakeholder":  r.who  || "",
        "날짜":          r.date || "",
        "요구사항":       r.text || "",
        "담당자":         r.assignee || "",
        "검토결과":       r.review   || "",
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
  // ID 컬럼처럼 긴 문자열은 Excel이 자동으로 숫자 변환할 수 있으므로
  // 셀 객체에 t:"s" 를 명시하면 안전하게 텍스트로 저장됨
  const aoaToSheet = (aoa) => {
    const ws = {};
    let maxC = 0;
    aoa.forEach((row, R) => {
      maxC = Math.max(maxC, row.length);
      row.forEach((cell, C) => {
        if (cell == null) return;
        const ref = XLSX.utils.encode_cell({r:R, c:C});
        if (typeof cell === "object" && "v" in cell) {
          // 셀 값이 _ 가 포함된 문자열(=ID 가능성)이면 강제 텍스트
          const v = cell.v;
          if (typeof v === "string" && /^[a-zA-Z]+_\d+/.test(v)) {
            ws[ref] = { ...cell, t: "s", v: String(v) };
          } else if (typeof v === "string") {
            ws[ref] = { ...cell, t: "s" };
          } else {
            ws[ref] = cell;
          }
        } else {
          // 원시값 셀: 문자열은 텍스트로
          if (typeof cell === "string") {
            ws[ref] = { v: cell, t: "s", s: XS.s({ h:"left" }) };
          } else {
            ws[ref] = { v: cell, s: XS.s({ h:"left" }) };
          }
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
    "비고","IF Type","ICD 번호","ICD 상태","TQ 번호","TQ 상태","Open Items","Edge ID",
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
    {wch:12},{wch:16},{wch:16},{wch:14},{wch:10},
    {wch:12},{wch:12},{wch:12},{wch:14},{wch:14},
    {wch:24},{wch:8},{wch:18},{wch:10},{wch:10},{wch:10},{wch:24},{wch:14},
  ];

  // ── SHEET 2: Equipment List ────────────────────────────────
  const eqHdrs = [
    "Item No.","설비명","설비 유형","소속 Area",
    "재질","용량","설계 압력","설계 온도","연결 Interface 수","비고","Node ID",
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
    {wch:12},{wch:14},{wch:12},{wch:12},{wch:16},{wch:28},{wch:14},
  ];

  // ── SHEET 3: Connection List ───────────────────────────────
  const cnHdrs = [
    "Line No.","Line Type","Fluid (Primary)","Fluid (Sub)",
    "Size","Schedule","Line Text",
    "From (Item No.)","To (Item No.)","연결 IC No.","IC 상태","Edge ID",
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
    {wch:16},{wch:16},{wch:12},{wch:12},{wch:14},
  ];

  // ── SHEET 4: Requirements ──────────────────────────────────
  const rqHdrs = [
    "Item No.","소속 Area","Stakeholder","날짜","요구사항","담당자","검토결과","Node ID",
  ];
  const rqAoa = [
    [Object.assign(XS.hdr("Requirements"), {
      s: XS.s({ bold:true, sz:12, bg:"1F3864", fc:"FFFFFF", bc:"1F3864" })
    }), ...Array(rqHdrs.length-1).fill(null)],
    rqHdrs.map(h => XS.shdr(h)),
    ...reqRows.map((row,i) => rqHdrs.map((h) => {
      const v = row[h] ?? "";
      const wrap = h==="요구사항" || h==="검토결과";
      return XS.alt(v, i,
        { h: (h==="요구사항"||h==="검토결과")?"left":"center", wrap,
          bold: h==="Item No.", fc: h==="Item No."?"1D4ED8":"000000" });
    })),
    ...(reqRows.length===0 ? [[XS.c("등록된 요구사항이 없습니다.",
      {h:"center",fc:"94A3B8",border:false})]] : []),
  ];
  const wsRq = aoaToSheet(rqAoa);
  wsRq["!merges"] = [{ s:{r:0,c:0}, e:{r:0,c:rqHdrs.length-1} }];
  wsRq["!rows"] = [{hpt:28},{hpt:22},...reqRows.map(()=>({hpt:36}))];
  wsRq["!cols"] = [
    {wch:14},{wch:22},{wch:16},{wch:12},{wch:42},{wch:14},{wch:32},{wch:14},
  ];

  // ── v10 SHEET: ICD Register ────────────────────────────────
  const icdHdrs = [
    "ICD 번호","ICD 상태","IF Type","Sender (Package)","Receiver (Package)",
    "IC 상태","IFA Date","IFC Date","TQ 번호","TQ 상태","Open Items",
  ];
  const icdAoa = [
    [Object.assign(XS.hdr("ICD Register — Interface Control Document"), {
      s: XS.s({ bold:true, sz:12, bg:"1F3864", fc:"FFFFFF", bc:"1F3864" })
    }), ...Array(icdHdrs.length-1).fill(null)],
    icdHdrs.map(h => XS.shdr(h)),
    ...icdRows.map((row,i) => icdHdrs.map((h) => {
      const v = row[h] ?? "";
      const wrap = h==="Open Items";
      return XS.alt(v, i,
        { h: h==="Open Items"?"left":"center", wrap,
          bold: h==="ICD 번호", fc: h==="ICD 번호"?"1D4ED8":"000000" });
    })),
    ...(icdRows.length===0 ? [[XS.c("등록된 ICD가 없습니다. (Edge의 ICD No. 입력 시 자동 집계)",
      {h:"center",fc:"94A3B8",border:false})]] : []),
  ];
  const wsICD = aoaToSheet(icdAoa);
  wsICD["!merges"] = [{ s:{r:0,c:0}, e:{r:0,c:icdHdrs.length-1} }];
  wsICD["!rows"] = [{hpt:28},{hpt:22},...icdRows.map(()=>({hpt:28}))];
  wsICD["!cols"] = [
    {wch:18},{wch:12},{wch:8},{wch:22},{wch:22},
    {wch:14},{wch:12},{wch:12},{wch:12},{wch:10},{wch:30},
  ];

  // ── v10.3 SHEET: IT Register ───────────────────────────────
  const itHdrs = [
    "IT No.","IC&TI Description","Source Area","Target Area",
    "공급 구분별 Inter-connection","기능 운영간 Inter-connection","Process 간 주요 영향인자",
    ...IT_CHECKLIST_CATEGORIES.flatMap(cat => [
      `${cat.label} (공급)`, `${cat.label} (공급담당)`,
      `${cat.label} (주관)`, `${cat.label} (주관담당)`,
      `${cat.label} (내용)`,
    ]),
    "Edge ID",
  ];
  const itAoa = [
    [Object.assign(XS.hdr("IT Register — Interface Tie-in (Area↔Area)"), {
      s: XS.s({ bold:true, sz:12, bg:"5B21B6", fc:"FFFFFF", bc:"5B21B6" })
    }), ...Array(itHdrs.length-1).fill(null)],
    itHdrs.map(h => XS.shdr(h)),
    ...itRows.map((row,i) => itHdrs.map((h) => {
      const v = row[h] ?? "";
      const wrap = h.includes("내용") || h.includes("기능 운영") ||
                   h.includes("영향인자") || h.includes("공급 구분별");
      return XS.alt(v, i,
        { h: wrap?"left":"center", wrap,
          bold: h==="IT No.", fc: h==="IT No."?"7C3AED":"000000" });
    })),
    ...(itRows.length===0 ? [[XS.c("등록된 IT Line이 없습니다. (Area↔Area 연결 시 자동 등록)",
      { italic:true, fc:"94A3B8", h:"center" }), ...Array(itHdrs.length-1).fill(null)]] : []),
  ];
  const wsIT = aoaToSheet(itAoa);
  wsIT["!merges"] = [
    { s:{r:0,c:0}, e:{r:0,c:itHdrs.length-1} },
    ...(itRows.length===0 ? [{ s:{r:2,c:0}, e:{r:2,c:itHdrs.length-1} }] : []),
  ];
  wsIT["!rows"] = [{hpt:28},{hpt:36},...itRows.map(()=>({hpt:60}))];
  wsIT["!cols"] = [
    {wch:8},{wch:32},{wch:18},{wch:18},
    {wch:30},{wch:36},{wch:30},
    ...IT_CHECKLIST_CATEGORIES.flatMap(()=>[{wch:8},{wch:12},{wch:10},{wch:12},{wch:32}]),
    {wch:24},
  ];

  // ── v10 SHEET: Scope Register ──────────────────────────────
  const scHdrs = [
    "WBS Code","Area","Type","Team","Discipline",
    "범위 기술","포함 범위","제외 범위",
    "Vendor 회사","Vendor 국가","계약번호","Interface Coordinator",
    "POSCO 담당","Engineering 담당",
  ];
  const scAoa = [
    [Object.assign(XS.hdr("Scope Register — Work Package & Vendor"), {
      s: XS.s({ bold:true, sz:12, bg:"1F3864", fc:"FFFFFF", bc:"1F3864" })
    }), ...Array(scHdrs.length-1).fill(null)],
    scHdrs.map(h => XS.shdr(h)),
    ...scopeRows.map((row,i) => scHdrs.map((h) => {
      const v = row[h] ?? "";
      const wrap = ["범위 기술","포함 범위","제외 범위","POSCO 담당","Engineering 담당"].includes(h);
      return XS.alt(v, i,
        { h: wrap?"left":"center", wrap,
          bold: h==="WBS Code", fc: h==="WBS Code"?"1D4ED8":"000000" });
    })),
    ...(scopeRows.length===0 ? [[XS.c("등록된 Scope가 없습니다. (Area 노드의 Scope 탭 입력 시 자동 집계)",
      {h:"center",fc:"94A3B8",border:false})]] : []),
  ];
  const wsSc = aoaToSheet(scAoa);
  wsSc["!merges"] = [{ s:{r:0,c:0}, e:{r:0,c:scHdrs.length-1} }];
  wsSc["!rows"] = [{hpt:28},{hpt:22},...scopeRows.map(()=>({hpt:40}))];
  wsSc["!cols"] = [
    {wch:16},{wch:18},{wch:10},{wch:8},{wch:12},
    {wch:30},{wch:24},{wch:24},
    {wch:16},{wch:12},{wch:14},{wch:18},
    {wch:28},{wch:32},
  ];

  // ── v10.1 SHEET 0: Scope of Supply ─────────────────────────
  const sosHdrs = [
    "구분","이름","Type","Vendor",
    ...SOS_POSCO_ROLES.map(r=>`POSCO·${r}`),
    ...SOS_SUPPLIER_ROLES.map(r=>`Supplier·${r}`),
    "Node ID",
  ];
  // 그룹 헤더 (POSCO / Supplier 구분)
  const poscoStart = 4;
  const poscoEnd   = poscoStart + SOS_POSCO_ROLES.length - 1;
  const supStart   = poscoEnd + 1;
  const supEnd     = supStart + SOS_SUPPLIER_ROLES.length - 1;
  const sosAoa = [
    // Row 0: 타이틀
    [Object.assign(XS.hdr("Scope of Supply — 담당자 배정표"), {
      s: XS.s({ bold:true, sz:13, bg:"1F3864", fc:"FFFFFF", bc:"1F3864" })
    }), ...Array(sosHdrs.length-1).fill(null)],
    // Row 1: 그룹 헤더 (기본정보 / POSCO / Supplier)
    [
      XS.c("기본 정보", {bold:true,bg:"D9E1F2",fc:"1F3864"}),
      null,null,null,
      XS.c("POSCO 담당", {bold:true,bg:"DDEBF7",fc:"1F3864"}),
      ...Array(SOS_POSCO_ROLES.length-1).fill(null),
      XS.c("Supplier 담당", {bold:true,bg:"FCE4D6",fc:"843C0C"}),
      ...Array(SOS_SUPPLIER_ROLES.length-1).fill(null),
      null,
    ],
    // Row 2: 컬럼 헤더
    sosHdrs.map(h => XS.shdr(h.replace("POSCO·","").replace("Supplier·",""))),
    // 데이터 행
    ...sosRows.map((row,i) => sosHdrs.map((h,j) => {
      const v = row[h] ?? "";
      const isPosco = h.startsWith("POSCO·");
      const isSup   = h.startsWith("Supplier·");
      return XS.alt(v, i, {
        h: j<1?"center":"left",
        bold: h==="이름"||h==="구분",
        fc: h==="구분" ? "1D4ED8"
           : isPosco ? "1F4E79"
           : isSup   ? "843C0C" : "000000",
      });
    })),
    ...(sosRows.length===0 ? [[XS.c("노드가 없습니다.",{h:"center",fc:"94A3B8",border:false})]] : []),
  ];
  const wsSOS = aoaToSheet(sosAoa);
  wsSOS["!merges"] = [
    { s:{r:0,c:0}, e:{r:0,c:sosHdrs.length-1} },             // 타이틀
    { s:{r:1,c:0}, e:{r:1,c:3} },                            // 기본정보
    { s:{r:1,c:poscoStart}, e:{r:1,c:poscoEnd} },            // POSCO
    { s:{r:1,c:supStart},   e:{r:1,c:supEnd} },              // Supplier
  ];
  wsSOS["!rows"] = [{hpt:30},{hpt:20},{hpt:22},...sosRows.map(()=>({hpt:22}))];
  wsSOS["!cols"] = [
    {wch:10},{wch:22},{wch:14},{wch:16},
    ...SOS_POSCO_ROLES.map(()=>({wch:14})),
    ...SOS_SUPPLIER_ROLES.map(()=>({wch:12})),
    {wch:14},
  ];

  XLSX.utils.book_append_sheet(wb, wsSOS, "Scope of Supply");
  XLSX.utils.book_append_sheet(wb, wsIC,  "IC Register");
  XLSX.utils.book_append_sheet(wb, wsICD, "ICD Register");
  XLSX.utils.book_append_sheet(wb, wsIT,  "IT Register");
  XLSX.utils.book_append_sheet(wb, wsEq,  "Equipment List");
  XLSX.utils.book_append_sheet(wb, wsCn,  "Connection List");
  XLSX.utils.book_append_sheet(wb, wsSc,  "Scope Register");
  XLSX.utils.book_append_sheet(wb, wsRq,  "Requirements");

  XLSX.writeFile(wb, `MBSE_ICRegister_${today}.xlsx`);
  return { icCount: icRows.length, equipCount: equipRows.length, connCount: connRows.length, icdCount: icdRows.length, scopeCount: scopeRows.length };
};

// ─────────────────────────────────────────────────────────────
// EXCEL IMPORT — IC Register → MBSE 모델 반영
// IC Register 시트: 상태·담당자·비고 → Edge data 업데이트
// Equipment List 시트: 스펙 → Node data 업데이트
// Requirements 시트: 요구사항 병합
// ─────────────────────────────────────────────────────────────
const importFromExcel = async (file, nodes, edges, setNodes, setEdges) => {
  // ═══════════════════════════════════════════════════════════
  // 진단 모드: 콘솔에 상세 로그를 출력합니다.
  // ═══════════════════════════════════════════════════════════
  console.group("%c[MBSE Import] 시작", "color:#1d4ed8;font-weight:bold;font-size:13px");
  console.log("📁 파일:", file.name, `(${(file.size/1024).toFixed(1)} KB)`);
  console.log("📊 현재 노드 수:", nodes.length, "/ 엣지 수:", edges.length);

  // 현재 노드/엣지의 ID 샘플
  console.log("🔍 현재 노드 ID 샘플 (앞 5개):",
    nodes.slice(0,5).map(n => ({id:n.id, type:n.type, label:n.data?.label||n.data?.itemNo||""})));
  console.log("🔍 현재 엣지 ID 샘플 (앞 5개):",
    edges.slice(0,5).map(e => ({id:e.id, src:e.source, tgt:e.target, ic:e.data?.ic_no||""})));

  const XLSX = await loadXLSX();
  console.log("✅ XLSX 라이브러리 로드 완료");

  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = ev => {
      try {
        const wb = XLSX.read(ev.target.result, { type:"array", cellStyles:false, cellFormula:false });
        console.log("📑 워크북에서 발견된 시트:", Object.keys(wb.Sheets));

        // ── DRM 암호화 파일 감지 ─────────────────────────────
        // DRM이 적용된 파일은 데이터가 알아볼 수 없는 바이너리로 들어옴
        // 첫 시트의 첫 셀에 "DRM", "encrypted", 또는 제어문자(\x00~\x1F)가 다수 포함되면 암호화로 판단
        const firstSheet = Object.values(wb.Sheets)[0];
        if (firstSheet) {
          const sample = XLSX.utils.sheet_to_json(firstSheet, { header:1, defval:null });
          const firstRow = sample[0] || [];
          const firstCellText = firstRow.map(c => {
            if (c == null) return "";
            const v = typeof c === "object" && "v" in c ? c.v : c;
            return String(v);
          }).join(" ");
          // DRM 패턴 검출
          const drmKeywords = ["DRM","encrypted","DRMONE","Fasoo","Markany","SoftCamp"];
          const hasDRMKeyword = drmKeywords.some(k => firstCellText.includes(k));
          // 제어문자 비율 (정상 엑셀은 거의 0%)
          const ctrlChars = (firstCellText.match(/[\x00-\x08\x0E-\x1F]/g)||[]).length;
          const ctrlRatio = firstCellText.length > 0 ? ctrlChars / firstCellText.length : 0;

          if (hasDRMKeyword || ctrlRatio > 0.1) {
            console.error("%c❌ DRM 암호화 파일 감지!", "color:#dc2626;font-weight:bold;font-size:14px");
            console.error("이 파일은 회사 보안 정책(DRM)에 의해 암호화되어 있어 import할 수 없습니다.");
            console.error("해결 방법:");
            console.error("  1. Excel에서 파일을 연 뒤 '다른 이름으로 저장' → 외부 폴더(예: 데스크탑)에 저장");
            console.error("  2. 또는 새 엑셀 파일에 데이터를 복사·붙여넣기 후 외부 폴더에 저장");
            console.error("  3. IT 보안팀에 DRM 예외 신청");
            console.groupEnd();
            const err = new Error("DRM 암호화 파일입니다. 외부 폴더에 새로 저장한 파일을 사용해주세요.");
            err.code = "DRM_DETECTED";
            reject(err);
            return;
          }
        }

        let updatedNodes = nodes.map(n => ({...n, data:{...(n.data||{})}}));
        let updatedEdges = edges.map(e => ({...e, data:{...(e.data||{})}}));
        const log = [];

        // ── 유틸: 셀 값 정규화 ─────────────────────────────────
        // sheet_to_json이 가끔 {v: 값, t: ...} 객체를 반환할 수 있고,
        // 특히 cellStyles=true 일 때 styled cell 객체가 반환됨
        const cellVal = (v) => {
          if (v == null) return "";
          if (typeof v === "object" && "v" in v) return String(v.v ?? "");
          return String(v);
        };
        // 셀이 시트에 있고 값이 비어있지 않으면 true (덮어쓰기 판단용)
        const hasVal = (row, key) => {
          if (!(key in row)) return false;
          const raw = row[key];
          if (raw == null) return false;
          if (typeof raw === "object" && "v" in raw) return raw.v != null;
          return true;
        };
        // 안전한 값 추출: 키 없으면 fallback 반환
        const pick = (row, key, fallback) => {
          if (!(key in row)) return fallback;
          const raw = row[key];
          if (raw == null) return "";
          if (typeof raw === "object" && "v" in raw) return raw.v == null ? "" : String(raw.v);
          return String(raw);
        };

        // ── 유연한 ID 매칭 (Excel이 긴 ID prefix를 잘라먹는 경우 대응) ──
        // 예) "area_1777534685748_10" → Excel "777534685748_10"
        //     "eq_1777537399751_2"   → Excel "777537399751_2"
        const findNodeByIdLoose = (idRaw, fallbackItemNo, fallbackLabel) => {
          const idStr = String(idRaw||"").trim();
          // 1) 정확 매칭
          if (idStr) {
            let n = updatedNodes.find(x => x.id === idStr);
            if (n) return n;
            // 2) suffix 매칭 (양방향)
            n = updatedNodes.find(x => x.id.endsWith(idStr) || idStr.endsWith(x.id));
            if (n) return n;
            // 3) 마지막 "_숫자_숫자" 패턴 추출 후 매칭
            const tail = idStr.match(/(\d+_\d+)$/)?.[1];
            if (tail) {
              n = updatedNodes.find(x => x.id.endsWith(tail));
              if (n) return n;
            }
          }
          // 4) Item No. fallback
          const itm = String(fallbackItemNo||"").trim();
          if (itm) {
            const n = updatedNodes.find(x => x.data?.itemNo === itm);
            if (n) return n;
          }
          // 5) Label fallback (Area)
          const lbl = String(fallbackLabel||"").trim();
          if (lbl) {
            const n = updatedNodes.find(x => x.type==="area" && x.data?.label === lbl);
            if (n) return n;
          }
          return null;
        };

        const findEdgeByIdLoose = (idRaw, fallbackICNo) => {
          const idStr = String(idRaw||"").trim();
          if (idStr) {
            let e = updatedEdges.find(x => x.id === idStr);
            if (e) return e;
            e = updatedEdges.find(x => x.id.endsWith(idStr) || idStr.endsWith(x.id));
            if (e) return e;
            const tail = idStr.match(/(\d+_\d+)$/)?.[1];
            if (tail) {
              e = updatedEdges.find(x => x.id.endsWith(tail));
              if (e) return e;
            }
          }
          // IC No. fallback
          const icNo = String(fallbackICNo||"").trim();
          if (icNo) {
            const e = updatedEdges.find(x => x.data?.ic_no === icNo);
            if (e) return e;
          }
          return null;
        };

        // ── 시트 찾기 (이름 변형 허용 + 컬럼명 fallback) ──────
        // 시트명이 안 맞아도 헤더 컬럼 구성으로 시트 종류 자동 식별
        const findSheet = (...candidates) => {
          const names = Object.keys(wb.Sheets);
          // 1) 이름 정확/유사 매칭
          for (const cand of candidates) {
            if (wb.Sheets[cand]) return wb.Sheets[cand];
            const norm = s => String(s).toLowerCase().replace(/[\s_\-]/g, "");
            const target = norm(cand);
            const hit = names.find(n => norm(n) === target);
            if (hit) return wb.Sheets[hit];
          }
          return null;
        };

        // ── 컬럼명 기반 시트 자동 식별 (fallback) ─────────────
        // 시트가 1개로 합쳐졌거나 이름이 완전히 바뀐 경우 사용
        // 시트 안의 헤더 컬럼명을 보고 어떤 종류 데이터인지 판별
        const findSheetByColumns = (signatureCols) => {
          for (const sheetName of Object.keys(wb.Sheets)) {
            const ws = wb.Sheets[sheetName];
            const raw = XLSX.utils.sheet_to_json(ws, { header:1, defval:null, blankrows:false });
            // 첫 15행 안에 signature 컬럼이 모두 들어있는 행을 찾음
            for (let i = 0; i < Math.min(raw.length, 15); i++) {
              const r = raw[i];
              if (!Array.isArray(r)) continue;
              const cellTexts = r.map(c => {
                if (c == null) return "";
                const v = typeof c === "object" && "v" in c ? c.v : c;
                return String(v).trim();
              });
              // 모든 signature 컬럼이 이 행에 존재하면 매칭
              const allFound = signatureCols.every(col => cellTexts.includes(col));
              if (allFound) {
                console.log(`  🔎 컬럼 매칭: 시트 "${sheetName}" 가 ${signatureCols.join(",")} 헤더 보유 → 사용`);
                return ws;
              }
            }
          }
          return null;
        };

        // ── 헬퍼: 스타일 적용 시트의 헤더 행 자동 감지 ───────
        // keyCol은 단일 문자열 또는 후보 배열. 후보 중 하나라도 발견되면 그 행이 헤더.
        const readSheet = (ws, keyCol, sheetLabel) => {
          if (!ws) { console.log(`  ⚠ ${sheetLabel}: 시트 없음`); return []; }
          const keyCandidates = Array.isArray(keyCol) ? keyCol : [keyCol];
          // 1) raw로 읽어서 헤더 행 위치 찾기 (defval=null로 빈 칸도 포함)
          const raw = XLSX.utils.sheet_to_json(ws, { header:1, defval:null, blankrows:false });
          let headerRow = 0;
          for (let i = 0; i < Math.min(raw.length, 15); i++) {
            const r = raw[i];
            if (!r || !Array.isArray(r)) continue;
            const cellTexts = r.map(c => {
              if (c == null) return "";
              const v = typeof c === "object" && "v" in c ? c.v : c;
              return String(v).trim();
            });
            // 후보 중 하나라도 이 행에 있으면 헤더로 판정
            const found = keyCandidates.some(k => cellTexts.includes(k));
            if (found) { headerRow = i; break; }
          }
          // 2) range로 객체 배열 추출 (defval:""로 빈 칸도 키 보장)
          const rows = XLSX.utils.sheet_to_json(ws, { range: headerRow, defval: "" });
          // 진단 로그
          console.log(`  📋 ${sheetLabel}: headerRow=${headerRow}, ${rows.length}행 발견`);
          if (rows.length > 0) {
            console.log(`     컬럼:`, Object.keys(rows[0]));
            console.log(`     첫 행:`, rows[0]);
          }
          return rows;
        };

        // ── IC Register 시트 → Edge 업데이트 ─────────────
        const wsIC = findSheet("IC Register","ICRegister","IC_Register")
                  || findSheetByColumns(["IC No.","상태","Edge ID"]);
        if (wsIC) {
          const rows = readSheet(wsIC, "Edge ID", "IC Register");
          let matched = 0, unmatched = 0;
          rows.forEach(row => {
            const edgeId = pick(row, "Edge ID", "").trim();
            const icNo = pick(row, "IC No.", "");
            const edge = findEdgeByIdLoose(edgeId, icNo);
            if (!edge) { if (edgeId||icNo) unmatched++; return; }
            const idx = updatedEdges.findIndex(e => e.id === edge.id);
            const e = updatedEdges[idx];
            const newData = { ...e.data };
            const apply = (key, dataKey) => {
              // v10.9: 빈 셀은 기존 값 유지 (엑셀 빈 칸이 앱 데이터를 지우는 것 방지)
              if (!(key in row)) return;
              const v = pick(row, key, "");
              if (String(v).trim() === "") return;
              newData[dataKey] = v;
            };
            apply("Line No.",       "serialNo");
            apply("Schedule",       "spec");
            apply("Line Text",      "lineText");
            apply("IC No.",         "ic_no");
            apply("상태",            "ic_status");
            apply("우선순위",        "ic_priority");
            apply("목표 완료일",     "ic_due");
            apply("실제 완료일",     "ic_closed");
            apply("담당자 (From)",   "ic_resp_from");
            apply("담당자 (To)",     "ic_resp_to");
            apply("비고",            "ic_remark");
            apply("ICD 번호",        "icd_no");
            apply("IF Type",         "ifType");
            apply("ICD 상태",        "icd_status");
            apply("TQ 번호",         "tq_no");
            apply("TQ 상태",         "tq_status");
            apply("Open Items",      "openItems");
            updatedEdges[idx] = { ...e, data: newData };
            matched++;
          });
          if (matched) log.push(`IC Register: ${matched}건 반영${unmatched?`, ${unmatched}건 미매칭`:""}`);
        }

        // ── Equipment List 시트 → Node 업데이트 ──────────
        const wsEq = findSheet("Equipment List","EquipmentList","Equipment")
                  || findSheetByColumns(["Item No.","설비명","설비 유형","Node ID"]);
        if (wsEq) {
          const rows = readSheet(wsEq, "Node ID", "Equipment List");
          let matched = 0, unmatched = 0;
          rows.forEach(row => {
            const idRaw  = pick(row, "Node ID", "");
            const itemNo = pick(row, "Item No.", "");
            const areaLbl= pick(row, "소속 Area", "");
            const node = findNodeByIdLoose(idRaw, itemNo, areaLbl);
            if (!node) { if (idRaw||itemNo) unmatched++; return; }
            const idx = updatedNodes.findIndex(n => n.id === node.id);
            const n = updatedNodes[idx];
            const newData = { ...n.data };
            const apply = (key, dataKey) => {
              // v10.9: 빈 셀은 기존 값 유지 (엑셀 빈 칸이 앱 데이터를 지우는 것 방지)
              if (!(key in row)) return;
              const v = pick(row, key, "");
              if (String(v).trim() === "") return;
              newData[dataKey] = v;
            };
            apply("Item No.",   "itemNo");
            apply("설비명",      "label");
            apply("재질",        "material");
            apply("용량",        "capacity");
            apply("설계 압력",   "designP");
            apply("설계 온도",   "designT");
            apply("비고",        "summary");
            updatedNodes[idx] = { ...n, data: newData };
            matched++;
          });
          if (matched) log.push(`Equipment List: ${matched}건 반영${unmatched?`, ${unmatched}건 미매칭`:""}`);
        }

        // ── Connection List 시트 → Edge 업데이트 ─────────
        const wsConn = findSheet("Connection List","ConnectionList","Connection")
                    || findSheetByColumns(["Line Type","Fluid (Primary)","Edge ID"]);
        if (wsConn) {
          const rows = readSheet(wsConn, "Edge ID", "Connection List");
          let matched = 0, unmatched = 0;
          rows.forEach(row => {
            const idRaw = pick(row, "Edge ID", "");
            const icNo  = pick(row, "연결 IC No.", "");
            const edge = findEdgeByIdLoose(idRaw, icNo);
            if (!edge) { if (idRaw) unmatched++; return; }
            const idx = updatedEdges.findIndex(e => e.id === edge.id);
            const e = updatedEdges[idx];
            const newData = { ...e.data };
            const apply = (key, dataKey) => {
              // v10.9: 빈 셀은 기존 값 유지 (엑셀 빈 칸이 앱 데이터를 지우는 것 방지)
              if (!(key in row)) return;
              const v = pick(row, key, "");
              if (String(v).trim() === "") return;
              newData[dataKey] = v;
            };
            apply("Line No.",         "serialNo");
            apply("Line Type",        "lineType");
            apply("Fluid (Primary)",  "fluidPrimary");
            apply("Fluid (Sub)",      "fluidSub");
            apply("Schedule",         "spec");
            apply("Line Text",        "lineText");
            if ("Size" in row) {
              const sz = pick(row, "Size", "");
              const m = sz.match(/^(\d+)A?$/);
              if (m) newData.sizeNum = m[1]; else newData.size = sz;
            }
            updatedEdges[idx] = { ...e, data: newData };
            matched++;
          });
          if (matched) log.push(`Connection List: ${matched}건 반영${unmatched?`, ${unmatched}건 미매칭`:""}`);
        }

        // ── Requirements 시트 → Node.requirements 업데이트 ──
        console.group("%c📝 Requirements 시트 처리", "color:#15803d;font-weight:bold");
        const wsR = findSheet("Requirements","Requirement")
                 || findSheetByColumns(["요구사항","Stakeholder","담당자"])
                 || findSheetByColumns(["요구사항","Stakeholder"])
                 || findSheetByColumns(["요구사항"]);

        if (!wsR) {
          // 진단: 모든 시트의 처음 5행을 덤프해서 어떤 컬럼이 들어있는지 확인
          console.warn("⚠ Requirements 시트 못 찾음. 모든 시트 내용 덤프:");
          Object.keys(wb.Sheets).forEach(name => {
            const ws = wb.Sheets[name];
            const raw = XLSX.utils.sheet_to_json(ws, { header:1, defval:null, blankrows:false });
            console.group(`📑 시트 "${name}" — 총 ${raw.length}행`);
            // 첫 8행 모두 출력
            raw.slice(0, 8).forEach((row, i) => {
              const cells = (row||[]).map(c => {
                if (c == null) return "";
                const v = typeof c === "object" && "v" in c ? c.v : c;
                return String(v).slice(0, 30);
              });
              console.log(`  Row ${i}:`, cells);
            });
            console.groupEnd();
          });
        }

        if (wsR) {
          const rows = readSheet(wsR, ["Node ID","Item No.","요구사항","Stakeholder"], "Requirements");
          const reqMap = {};
          let matched = 0, unmatched = 0;
          const unmatchedDetails = [];
          rows.forEach((r, rowIdx) => {
            const idRaw    = pick(r, "Node ID", "");
            const itemNo   = pick(r, "Item No.", "");
            const areaLbl  = pick(r, "소속 Area", "");
            const reqText  = pick(r, "요구사항", "");
            const assignee = pick(r, "담당자", "");
            const review   = pick(r, "검토결과", "");

            if (!reqText) {
              console.log(`  [행 ${rowIdx+1}] 요구사항 비어있음 → 스킵`);
              return;
            }

            const node = findNodeByIdLoose(idRaw, itemNo, areaLbl);
            if (!node) {
              unmatched++;
              unmatchedDetails.push({행:rowIdx+1, "Node ID":idRaw, "Item No.":itemNo, "소속 Area":areaLbl, 요구사항:reqText.slice(0,30)});
              return;
            }

            console.log(`  [행 ${rowIdx+1}] ✅ Node ID="${idRaw}" → 매칭된 노드: ${node.id} (${node.data?.label||node.data?.itemNo||""})`);
            console.log(`            담당자="${assignee}", 검토결과="${review}"`);

            if (!reqMap[node.id]) reqMap[node.id] = [];
            reqMap[node.id].push({
              id:   Date.now() + Math.random(),
              text: reqText,
              who:  pick(r, "Stakeholder", ""),
              date: pick(r, "날짜", ""),
              assignee: assignee,
              review:   review,
            });
            matched++;
          });

          if (unmatchedDetails.length > 0) {
            console.group("❌ 미매칭 행 상세:");
            console.table(unmatchedDetails);
            console.groupEnd();
          }

          updatedNodes = updatedNodes.map(n => {
            if (!reqMap[n.id]) return n;
            return { ...n, data:{ ...n.data, requirements: reqMap[n.id] } };
          });

          console.log(`  📊 결과: ${matched}건 매칭, ${unmatched}건 미매칭`);
          if (matched > 0) log.push(`Requirements: ${matched}건 반영${unmatched?`, ${unmatched}건 미매칭`:""}`);
          else if (unmatched > 0) log.push(`Requirements: ${unmatched}건 모두 미매칭 (Node ID 확인 필요)`);
        } else {
          console.log("  ⚠ Requirements 시트가 없습니다");
        }
        console.groupEnd();

        // ── Scope Register 시트 → Area Node 업데이트 ──
        const wsSc = findSheet("Scope Register","ScopeRegister","Scope")
                  || findSheetByColumns(["WBS Code","Vendor 회사","Node ID"]);
        if (wsSc) {
          const rows = readSheet(wsSc, "Node ID", "Scope Register");
          let matched = 0, unmatched = 0;
          rows.forEach(row => {
            const idRaw = pick(row, "Node ID", "");
            const areaLbl = pick(row, "Area", "");
            const node = findNodeByIdLoose(idRaw, "", areaLbl);
            if (!node) { if (idRaw) unmatched++; return; }
            const idx = updatedNodes.findIndex(n => n.id === node.id);
            const n = updatedNodes[idx];
            const parseStaff = (str) => {
              const obj = {};
              String(str||"").split("/").forEach(part => {
                const m = part.split(":");
                if (m.length===2) obj[m[0].trim()] = m[1].trim();
              });
              return obj;
            };
            const newData = { ...n.data };
            const apply = (key, dataKey) => {
              // v10.9: 빈 셀은 기존 값 유지 (엑셀 빈 칸이 앱 데이터를 지우는 것 방지)
              if (!(key in row)) return;
              const v = pick(row, key, "");
              if (String(v).trim() === "") return;
              newData[dataKey] = v;
            };
            apply("WBS Code", "wbsCode");
            apply("Team", "teamId");
            apply("Discipline", "discipline");
            apply("범위 기술", "scopeText");
            apply("포함 범위", "inclusions");
            apply("제외 범위", "exclusions");
            apply("Vendor 회사", "vendorName");
            apply("Vendor 국가", "vendorCountry");
            apply("계약번호", "vendorContract");
            apply("Interface Coordinator", "vendorICA");
            if ("POSCO 담당" in row && pick(row,"POSCO 담당","")) newData.poscoStaff = parseStaff(pick(row,"POSCO 담당",""));
            if ("Engineering 담당" in row && pick(row,"Engineering 담당","")) newData.engStaff = parseStaff(pick(row,"Engineering 담당",""));
            updatedNodes[idx] = { ...n, data: newData };
            matched++;
          });
          if (matched) log.push(`Scope Register: ${matched}건 반영${unmatched?`, ${unmatched}건 미매칭`:""}`);
        }

        // ── Scope of Supply 시트 → Node sos 업데이트 ──
        const wsSOS = findSheet("Scope of Supply","ScopeOfSupply","SoS")
                   || findSheetByColumns(["POSCO·PM","Supplier·대표","Node ID"]);
        if (wsSOS) {
          const rows = readSheet(wsSOS, "Node ID", "Scope of Supply");
          let matched = 0, unmatched = 0;
          rows.forEach(row => {
            const idRaw = pick(row, "Node ID", "");
            const nameLbl = pick(row, "이름", "");
            const node = findNodeByIdLoose(idRaw, nameLbl, nameLbl);
            if (!node) { if (idRaw||nameLbl) unmatched++; return; }
            const idx = updatedNodes.findIndex(n => n.id === node.id);
            const n = updatedNodes[idx];
            const posco = {};
            SOS_POSCO_ROLES.forEach(r => {
              const v = pick(row, `POSCO·${r}`, "").trim();
              if (v) posco[r] = v;
            });
            const supplier = {};
            SOS_SUPPLIER_ROLES.forEach(r => {
              const v = pick(row, `Supplier·${r}`, "").trim();
              if (v) supplier[r] = v;
            });
            updatedNodes[idx] = {
              ...n,
              data: { ...n.data, sos: { posco, supplier } }
            };
            matched++;
          });
          if (matched) log.push(`Scope of Supply: ${matched}건 반영${unmatched?`, ${unmatched}건 미매칭`:""}`);
        }

        // ── v10.3: IT Register 시트 → itEdge 업데이트 ──
        const wsIT = findSheet("IT Register","ITRegister","IT_Register","IT")
                  || findSheetByColumns(["IT No.","IC&TI Description","Edge ID"])
                  || findSheetByColumns(["IT No.","Source Area","Target Area"]);
        if (wsIT) {
          const rows = readSheet(wsIT, ["Edge ID","IT No.","IC&TI Description"], "IT Register");
          let matched = 0, unmatched = 0;
          rows.forEach(row => {
            const edgeId = pick(row, "Edge ID", "").trim();
            const itNo   = pick(row, "IT No.", "").trim();
            // v10.9: 빈 행 가드 — edgeId/itNo 둘 다 없으면 skip
            //  ("".endsWith() 가 항상 true 라서 빈 행이 첫 IT 라인에
            //   오매칭되어 데이터를 빈 값으로 덮어쓰던 버그 수정)
            if (!edgeId && !itNo) return;
            // itEdge만 대상 — edgeId 있을 때만 suffix 매칭 허용
            const itCandidate = edgeId ? updatedEdges.find(e =>
              e.type === "itEdge" && (e.id === edgeId || edgeId.endsWith(e.id) || e.id.endsWith(edgeId))
            ) : null;
            // edgeId 매칭 실패 시 IT No. 로 fallback (비어있지 않을 때만)
            const edge = itCandidate || (itNo ? updatedEdges.find(e =>
              e.type === "itEdge" && (e.data?.itNo === itNo)
            ) : null);
            if (!edge) { unmatched++; return; }
            const idx = updatedEdges.findIndex(e => e.id === edge.id);
            const e = updatedEdges[idx];
            const newData = { ...e.data };
            const apply = (key, dataKey) => {
              // v10.9: 빈 셀은 기존 값 유지 (엑셀 빈 칸이 앱 데이터를 지우는 것 방지)
              if (!(key in row)) return;
              const v = pick(row, key, "");
              if (String(v).trim() === "") return;
              newData[dataKey] = v;
            };
            apply("IT No.", "itNo");
            apply("기능 운영간 Inter-connection", "functionalDesc");

            // 공급 구분별 → "AreaA: 공급처A / AreaB: 공급처B" 파싱
            if ("공급 구분별 Inter-connection" in row) {
              const str = pick(row, "공급 구분별 Inter-connection", "");
              const parts = str.split("/").map(s => s.trim()).filter(Boolean);
              if (parts.length > 0) {
                newData.supplyByArea = parts.map(p => {
                  const m = p.match(/^(.+?):\s*(.+)$/);
                  return m ? { areaName: m[1].trim(), supplier: m[2].trim() }
                           : { areaName: p, supplier: "" };
                });
              }
            }
            // 영향인자 → "; " 또는 "\n" split (빈 셀이면 기존 값 유지)
            if ("Process 간 주요 영향인자" in row) {
              const str = pick(row, "Process 간 주요 영향인자", "");
              if (str.trim() !== "") {
                // "내용 [진행현황|담당자]" 형식 역파싱 (메타 없으면 기본값)
                // 줄바꿈(\n, \r\n) 또는 세미콜론으로 구분 → 각각 별개 영향인자
                newData.influenceFactors = str.split(/[;\n\r]+/).map(s=>s.trim()).filter(Boolean)
                  .map(s => {
                    const m = s.match(/^(.*?)\s*\[([^\]]*)\]\s*$/);
                    if (!m) return { text: s, status: "확인 중", owner: "" };
                    const [st, ow] = m[2].split("|").map(x=>(x||"").trim());
                    return { text: m[1].trim(), status: st || "확인 중", owner: ow || "" };
                  });
              }
            }
            // 체크리스트 카테고리별 (빈 셀은 기존 값 유지)
            // v10.10: 신형식 정규화 + 담당자 컬럼 + 줄바꿈 → 개별 항목 파싱
            const cl = {};
            IT_CHECKLIST_CATEGORIES.forEach(cat => {
              const next = normalizeITChecklistCat(newData.checklist?.[cat.key]);
              const supV  = pick(row, `${cat.label} (공급)`, "");
              const supOV = pick(row, `${cat.label} (공급담당)`, "");
              const resV  = pick(row, `${cat.label} (주관)`, "");
              const resOV = pick(row, `${cat.label} (주관담당)`, "");
              const cntV  = pick(row, `${cat.label} (내용)`, "");
              if (supV.trim()  !== "") next.supplier         = supV;
              if (supOV.trim() !== "") next.supplierOwner    = supOV;
              if (resV.trim()  !== "") next.responsible      = resV;
              if (resOV.trim() !== "") next.responsibleOwner = resOV;
              if (cntV.trim()  !== "") {
                // 줄바꿈(\n, \r\n) → 각 줄을 개별 항목으로, "내용 [진행현황]" 메타 파싱
                next.items = cntV.split(/[\n\r]+/).map(s=>s.trim()).filter(Boolean)
                  .map(s => {
                    const m = s.match(/^(.*?)\s*\[([^\]]*)\]\s*$/);
                    if (!m) return { text: s, status: "확인 중" };
                    return { text: m[1].trim(), status: (m[2]||"").trim() || "확인 중" };
                  });
              }
              cl[cat.key] = next;
            });
            newData.checklist = cl;

            updatedEdges[idx] = { ...e, data: newData };
            matched++;
          });
          if (matched) log.push(`IT Register: ${matched}건 반영${unmatched?`, ${unmatched}건 미매칭`:""}`);
        }

        // 강제 새 배열 (React 변경 감지 보장) + zIndex 정규화
        const finalNodes = normalizeAreaZIndex([...updatedNodes]);
        const finalEdges = [...updatedEdges];

        console.group("%c💾 최종 적용 직전 상태 확인", "color:#7c3aed;font-weight:bold");
        console.log(`업데이트된 노드 수: ${finalNodes.length}, 엣지 수: ${finalEdges.length}`);
        // 요구사항이 들어간 노드 샘플
        const nodesWithReq = finalNodes.filter(n => (n.data?.requirements||[]).length > 0);
        console.log(`Requirements 보유 노드: ${nodesWithReq.length}개`);
        if (nodesWithReq.length > 0) {
          const sample = nodesWithReq[0];
          console.log(`  샘플 노드 [${sample.id}] 요구사항:`,
            sample.data.requirements.map(r => ({
              text: r.text?.slice(0,40),
              assignee: r.assignee,
              review: r.review?.slice(0,40),
            })));
        }
        console.groupEnd();

        setNodes(finalNodes);
        setEdges(finalEdges);

        const msg = log.length > 0
          ? `Import 완료 — ${log.join(", ")}`
          : "Import 처리됨 (반영된 변경 없음)";

        console.log("%c[MBSE Import] 결과 요약", "color:#1d4ed8;font-weight:bold");
        console.log(msg);
        console.log("📋 시트별 로그:", log);
        console.groupEnd(); // 메인 그룹 닫기
        resolve({ log, msg });
      } catch(err) {
        console.error("❌ [MBSE Import] 오류:", err);
        console.groupEnd();
        reject(err);
      }
    };
    reader.onerror = (e) => {
      console.error("❌ [MBSE Import] 파일 읽기 오류:", e);
      console.groupEnd();
      reject(e);
    };
    reader.readAsArrayBuffer(file);
  });
};


// ─────────────────────────────────────────────────────────────
// GLOBAL CSS
// ─────────────────────────────────────────────────────────────
const GLOBAL_CSS = `
  /* ── Pipe 라인 호버 하이라이트 — 마우스 올리면 살짝 강조 ── */
  .mbse-pipe-edge:hover {
    filter: brightness(1.2) drop-shadow(0 0 2px rgba(37,99,235,0.35));
  }

  /* ─────────────────────────────────────────────
     IT (Interface Tie-in) 라인 — 평소 & 모드별 가시성
     ───────────────────────────────────────────── */
  /* 평소 — 얇고 흐리게 (P&ID 작업 화면 가리지 않도록) */
  .mbse-it-edge .mbse-it-edge-path {
    opacity: 0.4;
    transition: opacity 0.25s ease, stroke-width 0.25s ease;
  }
  .mbse-it-edge .mbse-it-edge-label {
    opacity: 0.6;
    transform-origin: center center;
    transition: opacity 0.25s ease, transform 0.25s ease;
  }
  /* 선택됐을 때 — 진한 강조 */
  .mbse-it-edge[data-selected="1"] .mbse-it-edge-path {
    opacity: 1;
  }
  .mbse-it-edge[data-selected="1"] .mbse-it-edge-label {
    opacity: 1;
  }
  .mbse-it-edge:hover .mbse-it-edge-path {
    opacity: 0.9;
  }
  .mbse-it-edge:hover .mbse-it-edge-label {
    opacity: 1;
  }

  /* Interface 모드 — IT 라인 강조, 다른 라인 페이드 */
  .react-flow.mbse-interface-mode .mbse-it-edge .mbse-it-edge-path {
    opacity: 1;
    stroke-width: 3.5px;
  }
  .react-flow.mbse-interface-mode .mbse-it-edge .mbse-it-edge-label {
    opacity: 1;
    /* Interface 모드 라벨 강조는 opacity 만 — transform 은 inline 에서 처리 */
  }
  /* Interface 모드에서 일반 라인 (pipe) 페이드 */
  .react-flow.mbse-interface-mode .react-flow__edges > svg > g.react-flow__edge[data-id]:not(:has(.mbse-it-edge)) {
    opacity: 0.12;
  }
  /* :has() 미지원 브라우저 fallback: edge 자체에 opacity */
  .react-flow.mbse-interface-mode .react-flow__edge {
    opacity: 0.12;
  }
  .react-flow.mbse-interface-mode .react-flow__edge:has(.mbse-it-edge),
  .react-flow.mbse-interface-mode .react-flow__edge.has-it-edge {
    opacity: 1;
  }
  /* 노드도 약간 톤다운 */
  .react-flow.mbse-interface-mode .react-flow__node-equipment,
  .react-flow.mbse-interface-mode .react-flow__node-instrument,
  .react-flow.mbse-interface-mode .react-flow__node-brench {
    opacity: 0.55;
  }

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

// v10.5: 줌 레벨 → 라벨 inverse scaling 계수 계산
// Zoom out 시 라벨이 자동으로 커지고, Zoom in 시 원래 크기로 작아짐
//   scale = clamp(1, 1/zoom, maxScale)
// 사용처: AreaNode 이름, IT 라인 라벨, IT 카운트 배지
const useZoomScale = (maxScale = 2.5) => {
  // ReactFlow store에서 현재 zoom 값 구독
  const zoom = useStore(s => s.transform?.[2] ?? 1);
  // zoom 이 1보다 작아질수록 scale 값 커짐, 1 이상이면 1 유지
  const raw  = zoom > 0 ? 1 / zoom : 1;
  return Math.min(Math.max(raw, 1), maxScale);
};

const AREA_TYPES = ["Plant","System","Package","Item"];
const AREA_COLORS = {
  Plant:    { bg:"rgba(191,219,254,0.65)", border:"#60a5fa", label:"#1e40af" },
  System:   { bg:"rgba(187,247,208,0.65)", border:"#4ade80", label:"#15803d" },
  Package:  { bg:"rgba(253,230,138,0.65)", border:"#facc15", label:"#92400e" },
  Item:     { bg:"rgba(233,213,255,0.65)", border:"#c084fc", label:"#6b21a8" },
};

// Area 노드 zIndex 정규화
// 큰 영역(Plant)은 항상 뒤, 작은 영역(Item)은 앞에 위치
// → 큰 Area를 선택해도 안에 있는 Equipment/하위 Area 클릭 가능
const AREA_ZINDEX = { Plant:-40, System:-30, Package:-20, Item:-10 };
const normalizeAreaZIndex = (nodes) => {
  if (!Array.isArray(nodes)) return nodes;
  return nodes.map(n => {
    if (n.type !== "area") return n;
    const z = AREA_ZINDEX[n.data?.areaType] ?? -10;
    return { ...n, zIndex: z };
  });
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

// ─────────────────────────────────────────────────────────────
// v10: Interface Type 분류 체계
// ─────────────────────────────────────────────────────────────
const IF_TYPES = {
  P:    { label:"P · Physical",    color:"#2563eb", desc:"배관·덕트·컨베이어 등 물리적 연결",  risk:"중"     },
  F:    { label:"F · Functional",  color:"#ea580c", desc:"제어신호·데이터·인터록",            risk:"높음"   },
  Perf: { label:"Perf · Performance", color:"#dc2626", desc:"성능 보증 경계 (유량·온도·압력)", risk:"매우높음" },
  T:    { label:"T · Temporal",    color:"#9333ea", desc:"운전 순서·시퀀스·타이밍",           risk:"높음"   },
};

// v10: IC Status 9단계 (외국사 분할발주 환경)
const IC_STATUS_FLOW = {
  "OPEN":          { color:"#ca8a04", next:"ICD 초안 작성 요청",   desc:"Interface 식별, 미협의" },
  "TQ-ISSUED":     { color:"#0891b2", next:"응답 대기 (기한 관리)", desc:"Vendor에 질의 발송" },
  "TQ-RESPONDED":  { color:"#0d9488", next:"기술 검토 및 수용/재질의", desc:"답변 수령" },
  "IFA":           { color:"#2563eb", next:"양사 검토 중",         desc:"합의 요청 단계" },
  "AGREED":        { color:"#7c3aed", next:"IFC 발행 대기",        desc:"양사 합의 완료" },
  "IFC":           { color:"#4f46e5", next:"설계 반영 확인",       desc:"시공용 확정" },
  "CLOSED":        { color:"#16a34a", next:"완결",                desc:"현장 검증 완료" },
  "DISPUTED":      { color:"#db2777", next:"발주처 중재 요청",     desc:"이견 발생" },
  "OVERDUE":       { color:"#dc2626", next:"에스컬레이션",         desc:"응답 기한 초과" },
};
const IC_STATUS_LIST = Object.keys(IC_STATUS_FLOW);

// v10: ICD Status (Interface Control Document)
const ICD_STATUS_LIST = ["IFD","IFA","IFC","Freeze","Superseded"];

// v10: TQ Status
const TQ_STATUS_LIST = ["발송됨","응답수령","추가질의","종결"];

// v10: Vendor 담당 조직 / 국가
const VENDOR_DISCIPLINES = ["PM","Civil","Architecture","Mechanical","Piping","EIC"];
const POSCO_ORGS = ["PM","운전","정비"];

// v10.1: Scope of Supply 담당자 역할 정의
const SOS_POSCO_ROLES    = ["PM","운전","정비"];                                       // 복수 입력
const SOS_SUPPLIER_ROLES = ["대표","PM","Mech.","Piping","Process","EIC","Civil","Arch"];

// ─────────────────────────────────────────────────────────────
// v10.3: IT (Interface Tie-in) — Area↔Area 통합 관리
//   캡쳐의 "Interconnection & Integration 정의" + "Check List" 양식 디지털화
// ─────────────────────────────────────────────────────────────
const IT_SUPPLIER_OPTIONS    = ["POSCO","ENC","DX","기타"];
const IT_RESPONSIBLE_OPTIONS = ["TF조업","TF설비","TF내화물","투엔실","ENC","DX","기타"];
const IT_CHECKLIST_CATEGORIES = [
  { key:"supply",  label:"장치 공급구분",      hdrBg:"#FEF3C7", hdrFc:"#92400E" },
  { key:"process", label:"Process",            hdrBg:"#DBEAFE", hdrFc:"#1E40AF" },
  { key:"mech",    label:"장치사양 (기계)",    hdrBg:"#FED7AA", hdrFc:"#9A3412" },
  { key:"piping",  label:"장치사양 (배관)",    hdrBg:"#CFFAFE", hdrFc:"#155E75" },
  { key:"refrac",  label:"장치사양 (내화물)",  hdrBg:"#FBCFE8", hdrFc:"#831843" },
  { key:"eic",     label:"장치사양 (EIC)",     hdrBg:"#BBF7D0", hdrFc:"#14532D" },
];
// v10.9: 영향인자 진행현황 옵션 + 상태 색상
const IT_FACTOR_STATUS = [
  { value:"확인 중", color:"#f59e0b" },
  { value:"완료",    color:"#16a34a" },
  { value:"변경 중", color:"#dc2626" },
];
// 영향인자 정규화: 구버전(문자열) → 신버전({text,status,owner}) 호환
const normalizeITFactor = (f) =>
  typeof f === "string"
    ? { text: f, status: "확인 중", owner: "" }
    : { text: f?.text || "", status: f?.status || "확인 중", owner: f?.owner || "" };

// v10.10: 체크리스트 카테고리 정규화
//  구버전 { supplier, responsible, content(문자열) }
//  → 신버전 { supplier, supplierOwner, responsible, responsibleOwner, items:[{text,status}] }
//  content 문자열은 줄바꿈 단위로 분해해 개별 항목으로 마이그레이션
const normalizeITChecklistCat = (c) => {
  if (!c) return { supplier:"ENC", supplierOwner:"", responsible:"TF조업", responsibleOwner:"", items:[] };
  let items;
  if (Array.isArray(c.items)) {
    items = c.items.map(it =>
      typeof it === "string"
        ? { text: it, status: "확인 중" }
        : { text: it?.text || "", status: it?.status || "확인 중" }
    );
  } else {
    items = String(c.content || "").split(/[\n\r]+/).map(s=>s.trim()).filter(Boolean)
      .map(t => ({ text: t, status: "확인 중" }));
  }
  return {
    supplier:         c.supplier || "",
    supplierOwner:    c.supplierOwner || "",
    responsible:      c.responsible || "",
    responsibleOwner: c.responsibleOwner || "",
    items,
  };
};
// IT 라인 default checklist 구조
const makeDefaultITChecklist = () => {
  const obj = {};
  IT_CHECKLIST_CATEGORIES.forEach(c => {
    obj[c.key] = { supplier:"ENC", supplierOwner:"", responsible:"TF조업", responsibleOwner:"", items:[] };
  });
  return obj;
};
// 카테고리에 작성된 항목이 있는지 (신·구 형식 호환)
const itCatHasContent = (c) =>
  (Array.isArray(c?.items) && c.items.some(it => (it?.text||"").trim()))
  || String(c?.content || "").trim().length > 0;


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
  const ax = areaNode.position.x, ay = areaNode.position.y;
  const aw = areaNode.style?.width  || areaNode.width  || 260;
  const ah = areaNode.style?.height || areaNode.height || 180;

  // 노드 bbox 가 Area 영역에 "겹치면" inside 로 판정
  // (좌상단만 검사하면 큰 노드 일부가 area 밖이어도 외부로 오판될 수 있음)
  const isInside = (n) => {
    if (n.id === areaNode.id) return false;
    const nx = n.position?.x ?? 0;
    const ny = n.position?.y ?? 0;
    const nw = n.measured?.width  ?? n.width  ?? n.style?.width  ?? (n.type==="brench"?24:120);
    const nh = n.measured?.height ?? n.height ?? n.style?.height ?? (n.type==="brench"?24:60);
    // 노드 중심점이 Area 영역 안에 있으면 inside
    const cx = nx + nw/2, cy = ny + nh/2;
    return cx >= ax && cy >= ay && cx <= ax + aw && cy <= ay + ah;
  };

  // 모든 노드(Area 포함 — 중첩 Area 내부의 노드도 따져야 함)에 대해
  // Area 안에 위치한 것들을 모은다. Equipment, Instrument, Brench, 그리고
  // 안에 포함된 하위 Area 까지 (단 inside 판정 자체에는 하위 Area 는 포함되지 않음 - 좌표만 확인)
  const insideIds = new Set();
  allNodes.forEach(n => {
    if (n.type === "area") return; // Area 자체는 노드로 카운트하지 않지만,
                                   // Area 내부의 노드들은 별도로 isInside 로 잡힌다
    if (isInside(n)) insideIds.add(n.id);
  });

  if (insideIds.size === 0) return { inlets:[], outlets:[] };

  const inlets = [], outlets = [];
  const seen = new Set();

  allEdges.forEach(e => {
    const si = insideIds.has(e.source);
    const ti = insideIds.has(e.target);
    if (si === ti) return; // 경계 통과 아님 (둘 다 안 or 둘 다 밖)

    const d = e.data || {};
    const sub  = d.fluidSub || d.lineType || "—";
    const size = d.size ? ` ${d.size}` : (d.sizeNum?` ${d.sizeNum}A`:"");
    const lineNo = d.serialNo ? ` ${d.serialNo}` : "";
    // 라벨에 라인 번호도 포함 → 같은 유체여도 별개 라인 구분 가능
    const label = `${sub}${size}${lineNo}`;

    // 외부 끝점(=영역 밖) 식별
    const outsideId = ti ? e.source : e.target;
    const outsideNode = allNodes.find(n => n.id === outsideId);
    // Brench면 한 단계 더 거슬러 올라가 진짜 외부 출처/목적지 찾기
    let realOutsideId = outsideId;
    if (outsideNode?.type === "brench") {
      realOutsideId = traceSource(outsideId, allNodes, allEdges) || outsideId;
    }
    const realOutsideNode = allNodes.find(n => n.id === realOutsideId);
    const outsideName =
      (realOutsideNode && realOutsideNode.type !== "brench")
        ? (realOutsideNode.data?.itemNo || realOutsideNode.data?.label || "외부")
        : "외부";

    // 중복 제거 키: 라벨 + 외부이름 + edge.id 까지 포함하여
    // "같은 유체라도 별개 라인"은 별도로 카운트
    const dir = ti ? "IN" : "OUT";
    const key = `${dir}:${e.id}`;          // edge id 기반 → 항상 유니크
    if (seen.has(key)) return;
    seen.add(key);

    const display = outsideName === "외부"
      ? label
      : `${label} (${dir==="IN"?"←":"→"} ${outsideName})`;

    if (ti) inlets.push(display);
    else    outlets.push(display);
  });

  return { inlets, outlets };
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

  // v10.5: Zoom out 시 라벨 자동 확대 (최대 2.5배)
  // Plant 처럼 큰 글자는 너무 크지 않도록 별도 한도, Item 은 살짝 더 크게
  const labelScale = useZoomScale(
    data.areaType === "Plant"   ? 2.0
    : data.areaType === "Item"  ? 2.5
    :                              2.2  // System / Package
  );

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

      {/* ── IT 전용 포트 2개 (상단 중앙, 하단 중앙) ──
          보라색 다이아몬드 모양으로 일반 포트와 시각적 차별화
          ID에 "it_" prefix → onConnect 에서 IT 라인으로 식별 */}
      {[
        { id:"it_top",    pos:Position.Top,    side:"top"    },
        { id:"it_bottom", pos:Position.Bottom, side:"bottom" },
      ].map(p => (
        <Handle key={p.id} type="source" position={p.pos} id={p.id}
          style={{
            width: 14, height: 14,
            background: "#7c3aed",
            border: "2px solid #fff",
            borderRadius: 2,
            transform: `translateX(-50%) rotate(45deg)`,
            left: "50%",
            boxShadow: "0 0 0 1px #7c3aed, 0 2px 6px rgba(124,58,237,0.45)",
            zIndex: 25,
            ...(p.side === "top"    ? { top:    -8 } : {}),
            ...(p.side === "bottom" ? { bottom: -8 } : {}),
          }}
          title="IT Line 전용 포트 (Interface Tie-in)"
        />
      ))}

      {/* ── IT 카운트 배지 (상단 중앙, 포트 위쪽) ──
          해당 Area에 연결된 IT 라인 수를 표시
          클릭 시 해당 Area의 IT 목록 팝업
          data.itCount는 App에서 edges 변경 시 자동 주입됨 */}
      {(data.itCount > 0) && (
        <div
          onClick={e => {
            e.stopPropagation();
            window.dispatchEvent(new CustomEvent("mbse:open-it-list", { detail:{ areaId: id }}));
          }}
          style={{
            position:"absolute",
            top: -28,
            left: "50%",
            // v10.5: 줌 아웃 시 IT 카운트 배지도 자동 확대 (최대 2.5배)
            transform: `translateX(-50%) scale(${labelScale})`,
            transformOrigin: "center bottom",
            transition: "transform 0.15s ease-out",
            background: "#7c3aed",
            color: "#fff",
            fontSize: 11,
            fontWeight: 800,
            padding: "3px 10px",
            borderRadius: 14,
            border: "2px solid #fff",
            boxShadow: "0 2px 6px rgba(124,58,237,0.4)",
            cursor: "pointer",
            zIndex: 30,
            whiteSpace: "nowrap",
            display: "flex",
            alignItems: "center",
            gap: 4,
            userSelect: "none",
          }}
          title={`이 Area의 IT 라인 ${data.itCount}개`}
        >
          🔗 {data.itCount}
        </div>
      )}

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
              // v10.5: 줌 아웃 시 라벨 자동 확대 (좌측 상단 기준)
              display:"inline-block",
              transform: `scale(${labelScale})`,
              transformOrigin: "left top",
              transition: "transform 0.15s ease-out",
            }}>
              {/* Plant/System/Package 는 태그 미표시, Item만 표시 */}
              {data.areaType==="Item" ? `[Item] ` : ""}
              {data.label || <span style={{ opacity:0.4, fontSize:"0.7em" }}>(더블클릭 편집)</span>}
            </span>
          )}
          {/* v10: Vendor 회사명 배지 */}
          {!editing && data.vendorName && (
            <div style={{ display:"flex",alignItems:"center",gap:3,marginTop:3 }}>
              <span style={{
                fontSize:9,fontWeight:700,background:"#1e293b",color:"#fff",
                padding:"0 6px",borderRadius:3,display:"inline-flex",alignItems:"center",gap:2,
              }}>🏭 {data.vendorName}{data.vendorCountry?` (${data.vendorCountry})`:""}</span>
              {data.teamId && (
                <span style={{
                  fontSize:9,fontWeight:700,
                  background:data.teamId==="FBR"?"#2563eb":data.teamId==="ESF"?"#dc2626":"#64748b",
                  color:"#fff",padding:"0 6px",borderRadius:3,
                }}>{data.teamId}</span>
              )}
            </div>
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

// ═══════════════════════════════════════════════════════════════
// P&ID-GRADE SMART ROUTING SYSTEM
//
// 적용 사양:
//  ① A* 그리드 기반 장애물 회피 (10px grid, 8px padding, cornerCost=5)
//  ② Strict orthogonal: 수평·수직만, 꺾임 radius=8px
//  ③ Multi-line parallel offset: 같은 segment 공유 시 4px 간격
//  ④ Live re-routing: 드래그 중엔 L-shape 빠른 경로
// ═══════════════════════════════════════════════════════════════

const ELBOW_R   = 8;    // 꺾임 반경 (P&ID 사양)
const MIN_STUB  = 20;   // 포트 이탈 최소 직선
const MARGIN    = 24;   // 장애물 여백 (드래그 중 빠른 라우팅용)
const SNAP_TOL  = 6;
const GRID      = 10;   // A* grid resolution
const PADDING   = 8;    // bbox 확장 padding
const CORNER_COST = 5;  // A* 꺾임 페널티
const PARALLEL_GAP = 10; // 같은 segment 공유 시 평행 간격 (P&ID 가독성)

// 좌표 → grid 스냅
const snapGrid = (v) => Math.round(v / GRID) * GRID;

// ── 포트 방향 벡터 ──────────────────────────────────────────
const portVec = pos => {
  if (pos==="right")  return {dx:1,dy:0};
  if (pos==="left")   return {dx:-1,dy:0};
  if (pos==="bottom") return {dx:0,dy:1};
  if (pos==="top")    return {dx:0,dy:-1};
  return {dx:1,dy:0};
};

// ═══════════════════════════════════════════════════════════════
// A* PATHFINDER — grid 기반 장애물 회피
// 입력: 시작/끝 좌표, 시작/끝 방향, 장애물 박스 배열
// 출력: orthogonal path 포인트 배열 (없으면 null)
// ═══════════════════════════════════════════════════════════════
const findPathAStar = (sx, sy, sDir, tx, ty, tDir, obstacles, bounds) => {
  // 시작점에서 포트 방향으로 충분히 이탈한 지점부터 탐색
  // (단 1 grid가 아닌 STUB_GRIDS 만큼 → Block 연결부 직선 구간 확보)
  const sv = portVec(sDir);
  const tv = portVec(tDir);
  const STUB_GRIDS = 2;   // 포트에서 2 grid(20px) 직선 이탈 후 탐색 시작
  const startX = sx + sv.dx * GRID * STUB_GRIDS;
  const startY = sy + sv.dy * GRID * STUB_GRIDS;
  const endX   = tx + tv.dx * GRID * STUB_GRIDS;
  const endY   = ty + tv.dy * GRID * STUB_GRIDS;

  // 좌표를 grid 단위로 변환
  const gSx = Math.round(startX / GRID), gSy = Math.round(startY / GRID);
  const gTx = Math.round(endX / GRID),   gTy = Math.round(endY / GRID);

  // 탐색 범위: src/tgt + 모든 장애물을 포함하도록 확장
  let minGX = Math.min(gSx, gTx);
  let maxGX = Math.max(gSx, gTx);
  let minGY = Math.min(gSy, gTy);
  let maxGY = Math.max(gSy, gTy);
  for (const o of obstacles) {
    minGX = Math.min(minGX, Math.floor(o.x1 / GRID));
    maxGX = Math.max(maxGX, Math.ceil(o.x2 / GRID));
    minGY = Math.min(minGY, Math.floor(o.y1 / GRID));
    maxGY = Math.max(maxGY, Math.ceil(o.y2 / GRID));
  }
  // 여유 추가
  minGX -= 8; maxGX += 8;
  minGY -= 8; maxGY += 8;

  // 장애물을 grid 단위로 미리 변환
  const gObs = obstacles.map(o => ({
    x1: Math.floor(o.x1 / GRID),
    y1: Math.floor(o.y1 / GRID),
    x2: Math.ceil(o.x2 / GRID),
    y2: Math.ceil(o.y2 / GRID),
  }));

  const isBlocked = (gx, gy) => {
    // 시작점과 끝점 grid는 항상 허용
    if ((gx===gSx && gy===gSy) || (gx===gTx && gy===gTy)) return false;
    for (const o of gObs) {
      if (gx >= o.x1 && gx <= o.x2 && gy >= o.y1 && gy <= o.y2) return true;
    }
    return false;
  };

  // direction: 0=right, 1=down, 2=left, 3=up
  const dirVec = [[1,0],[0,1],[-1,0],[0,-1]];
  const posToDir = (pos) => pos==="right"?0:pos==="bottom"?1:pos==="left"?2:3;

  const startDir = posToDir(sDir);

  // 우선순위 큐 (배열 기반)
  const heap = [];
  const visited = new Map();
  const key = (gx,gy,d) => `${gx},${gy},${d}`;

  heap.push({ gx:gSx, gy:gSy, dir:startDir, g:0, parent:null });

  let iter = 0;
  const MAX_ITER = 12000;

  while (heap.length > 0 && iter < MAX_ITER) {
    iter++;
    // 가장 낮은 f값 찾기
    let minIdx = 0;
    let minF = Infinity;
    for (let i = 0; i < heap.length; i++) {
      const c = heap[i];
      const h = Math.abs(c.gx-gTx) + Math.abs(c.gy-gTy);
      const f = c.g + h;
      if (f < minF) { minF = f; minIdx = i; }
    }
    const cur = heap.splice(minIdx, 1)[0];

    // 도착
    if (cur.gx === gTx && cur.gy === gTy) {
      // 경로 복원: 시작점/끝점 stub 포함
      const path = [{ x: sx, y: sy }]; // 실제 src 포트
      // src stub 점 (포트 방향 직선 확보)
      path.push({ x: startX, y: startY });
      let n = cur;
      const grid = [];
      while (n) {
        grid.unshift({ x: n.gx * GRID, y: n.gy * GRID });
        n = n.parent;
      }
      // grid 경로 추가
      grid.forEach(p => path.push(p));
      // tgt stub 점 + 끝점
      path.push({ x: endX, y: endY });
      path.push({ x: tx, y: ty });
      return path;
    }

    const k = key(cur.gx, cur.gy, cur.dir);
    if (visited.has(k) && visited.get(k) <= cur.g) continue;
    visited.set(k, cur.g);

    // 4방향 탐색
    for (let d = 0; d < 4; d++) {
      const [dx, dy] = dirVec[d];
      const ngx = cur.gx + dx;
      const ngy = cur.gy + dy;
      if (ngx < minGX || ngx > maxGX || ngy < minGY || ngy > maxGY) continue;
      if (isBlocked(ngx, ngy)) continue;

      const turnCost = (cur.dir !== d) ? CORNER_COST : 0;
      const ng = cur.g + 1 + turnCost;

      const nk = key(ngx, ngy, d);
      if (visited.has(nk) && visited.get(nk) <= ng) continue;

      heap.push({ gx:ngx, gy:ngy, dir:d, g:ng, parent:cur });
    }
  }
  return null;
};

// ═══════════════════════════════════════════════════════════════
// MULTI-LINE PARALLEL OFFSET
// 같은 source-target 쌍 또는 같은 segment 공유 시 4px 평행 분산
// ═══════════════════════════════════════════════════════════════
const applyParallelOffset = (pts, edgeId, allEdges, source, target) => {
  // 같은 source-target 쌍의 edge들
  const siblings = allEdges.filter(e =>
    (e.source === source && e.target === target) ||
    (e.source === target && e.target === source)
  );
  if (siblings.length <= 1) return pts;

  const idx = siblings.findIndex(e => e.id === edgeId);
  if (idx === -1) return pts;

  // 가운데를 기준으로 ±offset
  const N = siblings.length;
  const offset = (idx - (N - 1) / 2) * PARALLEL_GAP;
  if (Math.abs(offset) < 0.5) return pts;

  // 첫/끝 점은 그대로, 중간 점들에 offset 적용
  // 수평 세그먼트 → y에 offset, 수직 세그먼트 → x에 offset
  if (pts.length < 3) return pts;
  const result = [pts[0]];
  for (let i = 1; i < pts.length - 1; i++) {
    const prev = pts[i-1], curr = pts[i], next = pts[i+1];
    // 이전 세그먼트가 수평이면 curr.y에 offset
    const prevHoriz = Math.abs(prev.y - curr.y) < 2;
    // 다음 세그먼트가 수평이면 curr도 같은 y 유지
    const nextHoriz = Math.abs(curr.y - next.y) < 2;
    let nx = curr.x, ny = curr.y;
    if (prevHoriz && !nextHoriz) ny = curr.y + offset;
    else if (!prevHoriz && nextHoriz) ny = curr.y + offset;
    if (!prevHoriz && nextHoriz) nx = curr.x + offset;
    else if (prevHoriz && !nextHoriz) nx = curr.x + offset;
    result.push({ x: nx, y: ny });
  }
  result.push(pts[pts.length-1]);
  return result;
};

// ── 경로 정규화 (스냅 + 중복 제거 + 직선 병합) ───────────
const normalizePath = pts => {
  if (!pts||pts.length<2) return pts||[];
  // Step1: 인접 좌표 스냅
  const s = pts.map(p=>({...p}));
  for (let i=1;i<s.length;i++) {
    const p=s[i-1],c=s[i];
    if (Math.abs(c.y-p.y)<=SNAP_TOL && Math.abs(c.x-p.x)>SNAP_TOL) s[i]={x:c.x,y:p.y};
    if (Math.abs(c.x-p.x)<=SNAP_TOL && Math.abs(c.y-p.y)>SNAP_TOL) s[i]={x:p.x,y:c.y};
  }
  // Step2: 중복 제거
  const d = s.filter((p,i)=>i===0||Math.hypot(p.x-s[i-1].x,p.y-s[i-1].y)>0.5);
  // Step3: 직선 병합
  const m=[d[0]];
  for (let i=1;i<d.length-1;i++){
    const a=m[m.length-1],c=d[i],n=d[i+1];
    if (!(Math.abs(a.y-c.y)<0.5&&Math.abs(c.y-n.y)<0.5)
      &&!(Math.abs(a.x-c.x)<0.5&&Math.abs(c.x-n.x)<0.5)) m.push(c);
  }
  m.push(d[d.length-1]);
  return m;
};

// ── 직각 강제 분해 (v10.8) ─────────────────────────────────
// 사선 세그먼트를 발견하면 중간점을 삽입해 L자로 분해
// → "모든 라인은 직각" 원칙을 어떤 경우에도 보장
//   (노드 이동·waypoint 조작 후 생기는 사선을 자동 복원)
// 끝점 인접 세그먼트는 포트 방향에 맞춰 분해 (자연스러운 진입/이탈)
const rectifyOrthogonal = (pts, srcPos, tgtPos) => {
  if (!pts || pts.length < 2) return pts || [];
  const sv = portVec(srcPos || "right");
  const tv = portVec(tgtPos || "left");
  const out = [pts[0]];
  for (let i = 1; i < pts.length; i++) {
    const a = out[out.length - 1], b = pts[i];
    const dx = Math.abs(b.x - a.x), dy = Math.abs(b.y - a.y);
    if (dx > 0.5 && dy > 0.5) {
      // 사선 발견 → L자 분해
      const isFirst = out.length === 1;            // source 포트 직후
      const isLast  = i === pts.length - 1;        // target 포트 직전
      if (isFirst) {
        // source 포트 방향 우선: 수평 포트면 수평 먼저, 수직 포트면 수직 먼저
        if (sv.dx !== 0) out.push({ x: b.x, y: a.y });
        else             out.push({ x: a.x, y: b.y });
      } else if (isLast) {
        // target 진입이 포트 방향과 일치하도록: 수평 포트면 마지막이 수평
        if (tv.dx !== 0) out.push({ x: a.x, y: b.y });
        else             out.push({ x: b.x, y: a.y });
      } else {
        out.push({ x: b.x, y: a.y }); // 중간: 수평 먼저 (일관성)
      }
    }
    out.push(b);
  }
  return normalizePath(out);
};

// ── 1 + 2번 요건: 기본 라우팅 + 모든 세그먼트에 중앙 waypoint ─
// MIN_STUB로 포트에서 직선 이탈 후 꺾임, 각 세그먼트 중앙에 wp 자동 삽입
const routePipe = (sx,sy,srcPos,tx,ty,tgtPos,storedWp) => {
  // 수동 waypoints 있으면 직각 강제 분해 + 정규화
  // (노드 이동으로 생긴 사선을 자동으로 L자 복원 → 직각 원칙 유지)
  if (storedWp&&storedWp.length>0)
    return rectifyOrthogonal([{x:sx,y:sy},...storedWp,{x:tx,y:ty}], srcPos, tgtPos);

  const sv=portVec(srcPos), tv=portVec(tgtPos);
  const stub = MIN_STUB; // 고정 최소 stub

  // stub 끝점
  const sx2=sx+sv.dx*stub, sy2=sy+sv.dy*stub;
  const tx2=tx+tv.dx*stub, ty2=ty+tv.dy*stub;

  let raw;
  // 수평 출발 ↔ 수평 도착
  if (sv.dy===0 && tv.dy===0) {
    if (Math.abs(sy-ty)<2) {
      raw=[{x:sx,y:sy},{x:tx,y:ty}];
    } else if ((sv.dx>0&&tx>sx)||(sv.dx<0&&tx<sx)) {
      const mx=(sx2+tx2)/2;
      raw=[{x:sx,y:sy},{x:sx2,y:sy2},{x:mx,y:sy2},{x:mx,y:ty2},{x:tx2,y:ty2},{x:tx,y:ty}];
    } else {
      const ux=sv.dx>0?Math.max(sx2,tx2)+stub:Math.min(sx2,tx2)-stub;
      raw=[{x:sx,y:sy},{x:sx2,y:sy2},{x:ux,y:sy2},{x:ux,y:ty2},{x:tx2,y:ty2},{x:tx,y:ty}];
    }
  }
  // 수직 출발 ↔ 수직 도착
  else if (sv.dx===0 && tv.dx===0) {
    if (Math.abs(sx-tx)<2) {
      raw=[{x:sx,y:sy},{x:tx,y:ty}];
    } else if ((sv.dy>0&&ty>sy)||(sv.dy<0&&ty<sy)) {
      const my=(sy2+ty2)/2;
      raw=[{x:sx,y:sy},{x:sx2,y:sy2},{x:sx2,y:my},{x:tx2,y:my},{x:tx2,y:ty2},{x:tx,y:ty}];
    } else {
      const uy=sv.dy>0?Math.max(sy2,ty2)+stub:Math.min(sy2,ty2)-stub;
      raw=[{x:sx,y:sy},{x:sx2,y:sy2},{x:sx2,y:uy},{x:tx2,y:uy},{x:tx2,y:ty2},{x:tx,y:ty}];
    }
  }
  // 수평 출발 → 수직 도착
  else if (sv.dy===0 && tv.dx===0) {
    raw=[{x:sx,y:sy},{x:sx2,y:sy2},{x:tx2,y:sy2},{x:tx2,y:ty2},{x:tx,y:ty}];
  }
  // 수직 출발 → 수평 도착
  else {
    raw=[{x:sx,y:sy},{x:sx2,y:sy2},{x:sx2,y:ty2},{x:tx2,y:ty2},{x:tx,y:ty}];
  }

  const norm = normalizePath(raw);

  // 2번 요건: 각 중간 세그먼트(stub 제외)에 중앙 waypoint 삽입
  if (norm.length < 2) return norm;
  const withMid = [norm[0]];
  for (let i=0;i<norm.length-1;i++){
    const a=norm[i],b=norm[i+1];
    const isFirst=i===0, isLast=i===norm.length-2;
    const len=Math.hypot(b.x-a.x,b.y-a.y);
    // stub(첫/마지막) 세그먼트는 중앙 wp 불필요
    if (!isFirst && !isLast && len>10) {
      withMid.push({x:(a.x+b.x)/2, y:(a.y+b.y)/2});
    }
    withMid.push(b);
  }
  return withMid;
};

// ── 3번 요건: 장애물 지그재그 회피 ──────────────────────────
// 세그먼트가 블록(equipment/brench/instrument)을 통과하면
// 블록 위 또는 아래(수평) / 왼쪽 또는 오른쪽(수직)으로 지그재그 우회
const avoidObstacles = (pts, obstacles) => {
  if (!obstacles||obstacles.length===0) return pts;

  // 선분과 박스 교차 여부
  const hits = (x1,y1,x2,y2,box) => {
    const minX=Math.min(x1,x2)-1, maxX=Math.max(x1,x2)+1;
    const minY=Math.min(y1,y2)-1, maxY=Math.max(y1,y2)+1;
    return !(maxX<box.x1||minX>box.x2||maxY<box.y1||minY>box.y2);
  };

  let res=[...pts];
  for (let iter=0;iter<6;iter++) {
    let changed=false;
    const next=[res[0]];
    for (let i=0;i<res.length-1;i++) {
      const a=res[i], b=res[i+1];
      const isH=Math.abs(a.y-b.y)<2;
      let hit=null;
      for (const box of obstacles) {
        if (hits(a.x,a.y,b.x,b.y,box)){hit=box;break;}
      }
      if (hit) {
        changed=true;
        if (isH) {
          // 수평 세그먼트 → 위/아래 중 src 쪽으로 더 가까운 방향
          const aboveY=hit.y1-MARGIN, belowY=hit.y2+MARGIN;
          const byY=Math.abs(a.y-aboveY)<=Math.abs(a.y-belowY)?aboveY:belowY;
          // 지그재그: a.x→byY, b.x→byY 사이에 우회 세그먼트 삽입
          next.push({x:a.x,y:byY},{x:b.x,y:byY});
        } else {
          // 수직 세그먼트 → 왼/오른 중 src 쪽으로 더 가까운 방향
          const leftX=hit.x1-MARGIN, rightX=hit.x2+MARGIN;
          const byX=Math.abs(a.x-leftX)<=Math.abs(a.x-rightX)?leftX:rightX;
          next.push({x:byX,y:a.y},{x:byX,y:b.y});
        }
      } else {
        next.push(b);
      }
    }
    if (!changed) break;
    res=normalizePath(next);
  }
  return res;
};

// ═══════════════════════════════════════════════════════════════
// v10.6 라우팅 품질 개선 헬퍼 (3대 문제 해결)
// ═══════════════════════════════════════════════════════════════

// ── 문제 1: 끝점 수직 stub 강제 ─────────────────────────────
// 포트(화살표)에서 핸들 방향으로 최소 stubLen 직선 확보
// → Block 연결부에서 화살표 수직 방향 직선 구간 보장
const PIPE_STUB = 24;   // Pipe 라인 끝점 최소 직선 (px)
const enforceEndpointStubs = (pts, srcPos, tgtPos, stubLen = PIPE_STUB) => {
  if (!pts || pts.length < 2) return pts;
  const sv = portVec(srcPos);
  const tv = portVec(tgtPos);
  const result = pts.map(p => ({ ...p }));

  // Source stub: 첫 세그먼트가 포트 방향으로 stubLen 이상 진행하는지 확인
  const s0 = result[0], s1 = result[1];
  const sProgress = (s1.x - s0.x) * sv.dx + (s1.y - s0.y) * sv.dy;
  // 포트 방향과 수직(엇나간) 성분
  const sPerp = Math.abs((s1.x - s0.x) * sv.dy) + Math.abs((s1.y - s0.y) * sv.dx);
  if (sProgress < stubLen - 0.5 || sPerp > 0.5) {
    // stub 점 강제 삽입 (포트 방향으로 stubLen)
    result.splice(1, 0, {
      x: s0.x + sv.dx * stubLen,
      y: s0.y + sv.dy * stubLen,
    });
  }

  // Target stub: 마지막 세그먼트
  const t0 = result[result.length - 1], t1 = result[result.length - 2];
  const tProgress = (t1.x - t0.x) * tv.dx + (t1.y - t0.y) * tv.dy;
  const tPerp = Math.abs((t1.x - t0.x) * tv.dy) + Math.abs((t1.y - t0.y) * tv.dx);
  if (tProgress < stubLen - 0.5 || tPerp > 0.5) {
    result.splice(result.length - 1, 0, {
      x: t0.x + tv.dx * stubLen,
      y: t0.y + tv.dy * stubLen,
    });
  }
  return normalizePath(result);
};

// ── 문제 2: 경로 단순화 ("산 모양" 펴기) ────────────────────
// A* grid 결과의 불필요한 계단식 꺾임 제거
// 같은 방향으로 갈 수 있는데 지그재그로 가는 구간을 직선화
const simplifyPath = (pts, obstacles) => {
  if (!pts || pts.length < 3) return pts;

  // 선분-박스 교차 검사 (장애물 통과 여부)
  const segHitsBox = (x1,y1,x2,y2,box) => {
    const minX=Math.min(x1,x2)-0.5, maxX=Math.max(x1,x2)+0.5;
    const minY=Math.min(y1,y2)-0.5, maxY=Math.max(y1,y2)+0.5;
    return !(maxX<box.x1||minX>box.x2||maxY<box.y1||minY>box.y2);
  };
  const pathClear = (a, b) => {
    if (!obstacles) return true;
    for (const box of obstacles) {
      if (segHitsBox(a.x, a.y, b.x, b.y, box)) return false;
    }
    return true;
  };

  let work = pts.map(p => ({ ...p }));

  // 반복: "ㄷ"자/계단형 꺾임을 L자로 단순화
  for (let pass = 0; pass < 4; pass++) {
    let changed = false;
    const out = [work[0]];
    let i = 0;
    while (i < work.length - 1) {
      const a = work[i];
      // a에서 갈 수 있는 가장 먼 점 j 찾기 (직각 유지 + 장애물 회피)
      let bestJ = i + 1;
      for (let j = i + 2; j < work.length; j++) {
        const c = work[j];
        // a→c 가 순수 수평 또는 수직이고 장애물 없으면 직선화 가능
        const isOrtho = Math.abs(a.x - c.x) < 1 || Math.abs(a.y - c.y) < 1;
        if (isOrtho && pathClear(a, c)) {
          bestJ = j;
          changed = true;
        }
      }
      out.push(work[bestJ]);
      i = bestJ;
    }
    work = normalizePath(out);
    if (!changed) break;
  }

  // "ㄷ"자 우회 단순화: 3점이 같은 축으로 되돌아가면 中점 제거 시도
  // (예: 오른쪽→위→오른쪽 보다 한 번에 가는 게 가능하면)
  for (let pass = 0; pass < 3; pass++) {
    if (work.length < 4) break;
    let changed = false;
    for (let i = 1; i < work.length - 2; i++) {
      const p0 = work[i-1], p1 = work[i], p2 = work[i+1], p3 = work[i+2];
      // p1-p2 세그먼트를 없애고 p0에서 p3로 L자로 갈 수 있나?
      // 후보 1: p0 → (p3.x, p0.y) → p3
      const cand1 = { x: p3.x, y: p0.y };
      if (pathClear(p0, cand1) && pathClear(cand1, p3)) {
        work.splice(i, 2, cand1);
        work = normalizePath(work);
        changed = true;
        break;
      }
      // 후보 2: p0 → (p0.x, p3.y) → p3
      const cand2 = { x: p0.x, y: p3.y };
      if (pathClear(p0, cand2) && pathClear(cand2, p3)) {
        work.splice(i, 2, cand2);
        work = normalizePath(work);
        changed = true;
        break;
      }
    }
    if (!changed) break;
  }

  return work;
};

// ── 문제 2 보조: 자동 경로에 조정 가능한 waypoint 삽입 ──────
// 중간 세그먼트 중앙에 waypoint를 넣어 사용자가 핸들로 잡을 수 있게 함
const insertMidWaypoints = (pts) => {
  if (!pts || pts.length < 2) return pts;
  const out = [pts[0]];
  for (let i = 0; i < pts.length - 1; i++) {
    const a = pts[i], b = pts[i+1];
    const isFirst = i === 0, isLast = i === pts.length - 2;
    const len = Math.hypot(b.x - a.x, b.y - a.y);
    if (!isFirst && !isLast && len > 16) {
      out.push({ x: (a.x + b.x) / 2, y: (a.y + b.y) / 2 });
    }
    out.push(b);
  }
  return out;
};

// ── 문제 3: 노드 이동 시 직각 보정 ──────────────────────────
// endpoint(포트)가 이동했을 때, 인접 세그먼트를 수평/수직으로 복원
// waypoint는 절대좌표라 endpoint만 따라가면 사선이 생기는 것을 방지
const reorthogonalizePath = (pts, srcPos, tgtPos) => {
  if (!pts || pts.length < 2) return pts;
  const sv = portVec(srcPos);
  const tv = portVec(tgtPos);
  const r = pts.map(p => ({ ...p }));

  // Source: 첫 waypoint를 포트 방향 축에 맞춤
  if (r.length >= 2) {
    const s0 = r[0], s1 = r[1];
    if (sv.dx !== 0) {
      // 수평 포트 → 첫 세그먼트는 수평이어야 함 → s1.y = s0.y
      r[1] = { ...s1, y: s0.y };
    } else {
      // 수직 포트 → 첫 세그먼트는 수직이어야 함 → s1.x = s0.x
      r[1] = { ...s1, x: s0.x };
    }
  }
  // Target: 마지막 waypoint를 포트 방향 축에 맞춤
  if (r.length >= 2) {
    const t0 = r[r.length-1], t1 = r[r.length-2];
    if (tv.dx !== 0) {
      r[r.length-2] = { ...t1, y: t0.y };
    } else {
      r[r.length-2] = { ...t1, x: t0.x };
    }
  }
  return normalizePath(r);
};

// ═══════════════════════════════════════════════════════════════
// v10.7 결정론적 직교 라우터 (draw.io / yEd 스타일)
//   · 모든 세그먼트 100% 수평/수직 (사선 금지)
//   · 포트에서 고정 stub 후 직각 진입
//   · L 또는 Z 형태의 단순 경로 (산 모양 방지)
//   · 장애물은 bounding box 단위로 우회
// ═══════════════════════════════════════════════════════════════
const ORTHO_STUB = 28;   // 포트 이탈 직선 길이 (px)
const ORTHO_MARGIN = 22; // 장애물 회피 여백 (px)

// 선분(수평 또는 수직)이 박스와 교차하는지
const segHitsBoxOrtho = (a, b, box) => {
  const minX = Math.min(a.x, b.x), maxX = Math.max(a.x, b.x);
  const minY = Math.min(a.y, b.y), maxY = Math.max(a.y, b.y);
  return !(maxX < box.x1 || minX > box.x2 || maxY < box.y1 || minY > box.y2);
};

// 경로 전체가 장애물과 충돌하는지
const pathClearOrtho = (pts, obstacles) => {
  if (!obstacles || obstacles.length === 0) return true;
  for (let i = 0; i < pts.length - 1; i++) {
    for (const box of obstacles) {
      if (segHitsBoxOrtho(pts[i], pts[i+1], box)) return false;
    }
  }
  return true;
};

// 메인 직교 라우터
const routeOrthogonal = (sx, sy, srcPos, tx, ty, tgtPos, obstacles) => {
  const sv = portVec(srcPos);
  const tv = portVec(tgtPos);

  // 1) 포트에서 stub 만큼 직각 이탈
  const sStub = { x: sx + sv.dx * ORTHO_STUB, y: sy + sv.dy * ORTHO_STUB };
  const tStub = { x: tx + tv.dx * ORTHO_STUB, y: ty + tv.dy * ORTHO_STUB };

  const src = { x: sx, y: sy };
  const tgt = { x: tx, y: ty };

  // 2) sStub → tStub 사이를 직각으로 연결하는 후보 경로들 생성
  //    모든 후보는 100% 수평/수직 세그먼트로만 구성
  const candidates = [];

  // 후보 A: sStub → (tStub.x, sStub.y) → tStub  (수평 먼저)
  candidates.push([src, sStub, { x: tStub.x, y: sStub.y }, tStub, tgt]);
  // 후보 B: sStub → (sStub.x, tStub.y) → tStub  (수직 먼저)
  candidates.push([src, sStub, { x: sStub.x, y: tStub.y }, tStub, tgt]);
  // 후보 C: Z형 수평 분할 (중간 X에서 꺾기)
  const midX = (sStub.x + tStub.x) / 2;
  candidates.push([src, sStub,
    { x: midX, y: sStub.y }, { x: midX, y: tStub.y },
    tStub, tgt]);
  // 후보 D: Z형 수직 분할 (중간 Y에서 꺾기)
  const midY = (sStub.y + tStub.y) / 2;
  candidates.push([src, sStub,
    { x: sStub.x, y: midY }, { x: tStub.x, y: midY },
    tStub, tgt]);

  // 3) 충돌 없는 후보들을 점수화 → 최적 선택
  //    점수 = 꺾임수×80 + 경로길이×0.5 + 역주행 페널티
  //    (꺾임이 적고 짧으며 포트 방향을 거스르지 않는 경로 우선)
  const scoreRoute = (norm) => {
    let bends = 0, len = 0;
    for (let i = 1; i < norm.length; i++) {
      len += Math.abs(norm[i].x - norm[i-1].x) + Math.abs(norm[i].y - norm[i-1].y);
      if (i < norm.length - 1) {
        const dx1 = norm[i].x - norm[i-1].x, dy1 = norm[i].y - norm[i-1].y;
        const dx2 = norm[i+1].x - norm[i].x, dy2 = norm[i+1].y - norm[i].y;
        if ((Math.abs(dx1) > 0.5) !== (Math.abs(dx2) > 0.5)) bends++;
      }
    }
    // 역주행 페널티: stub 직후 세그먼트가 포트 방향과 반대로 가면 +120
    let backtrack = 0;
    if (norm.length >= 3) {
      const p1 = norm[1], p2 = norm[2];
      const prog = (p2.x - p1.x) * sv.dx + (p2.y - p1.y) * sv.dy;
      if (prog < -0.5) backtrack += 120;
    }
    return bends * 80 + len * 0.5 + backtrack;
  };

  let best = null, bestScore = Infinity;
  for (const cand of candidates) {
    const norm = normalizePath(cand);
    if (pathClearOrtho(norm, obstacles)) {
      const s = scoreRoute(norm);
      if (s < bestScore) { bestScore = s; best = norm; }
    }
  }
  if (best) return best;

  // 4) 모든 단순 후보가 충돌 → A·B 기반 우회 경로를 각각 생성해 더 나은 쪽 선택
  const detA = detourAroundBoxes(normalizePath(candidates[0]), obstacles, sStub, tStub, src, tgt);
  const detB = detourAroundBoxes(normalizePath(candidates[1]), obstacles, sStub, tStub, src, tgt);
  const sA = scoreRoute(detA) + (pathClearOrtho(detA, obstacles) ? 0 : 500);
  const sB = scoreRoute(detB) + (pathClearOrtho(detB, obstacles) ? 0 : 500);
  return sA <= sB ? detA : detB;
};

// 박스 우회: 충돌 세그먼트를 박스 외곽으로 밀어냄
const detourAroundBoxes = (pts, obstacles, sStub, tStub, src, tgt) => {
  if (!obstacles || obstacles.length === 0) return pts;
  let work = pts.map(p => ({ ...p }));

  for (let iter = 0; iter < 8; iter++) {
    let hitFound = false;
    const out = [work[0]];

    for (let i = 0; i < work.length - 1; i++) {
      const a = work[i], b = work[i+1];
      const isH = Math.abs(a.y - b.y) < 1;

      // 이 세그먼트와 충돌하는 박스 찾기
      let box = null;
      for (const o of obstacles) {
        if (segHitsBoxOrtho(a, b, o)) { box = o; break; }
      }

      if (box) {
        hitFound = true;
        if (isH) {
          // 수평 세그먼트 → 박스 위 또는 아래로 우회
          const goAbove = Math.abs(a.y - (box.y1 - ORTHO_MARGIN)) <= Math.abs(a.y - (box.y2 + ORTHO_MARGIN));
          const detourY = goAbove ? box.y1 - ORTHO_MARGIN : box.y2 + ORTHO_MARGIN;
          out.push({ x: a.x, y: detourY });
          out.push({ x: b.x, y: detourY });
          out.push(b);
        } else {
          // 수직 세그먼트 → 박스 좌 또는 우로 우회
          const goLeft = Math.abs(a.x - (box.x1 - ORTHO_MARGIN)) <= Math.abs(a.x - (box.x2 + ORTHO_MARGIN));
          const detourX = goLeft ? box.x1 - ORTHO_MARGIN : box.x2 + ORTHO_MARGIN;
          out.push({ x: detourX, y: a.y });
          out.push({ x: detourX, y: b.y });
          out.push(b);
        }
      } else {
        out.push(b);
      }
    }

    work = normalizePath(out);
    if (!hitFound) break;
  }
  return work;
};

// ── SVG Rounded Elbow path ────────────────────────────────────
const buildElbowPath = (pts, r=ELBOW_R) => {
  if (!pts||pts.length<2) return "";
  const c=pts.filter((p,i)=>i===0||Math.hypot(p.x-pts[i-1].x,p.y-pts[i-1].y)>0.5);
  if (c.length<2) return "";
  let d=`M ${c[0].x} ${c[0].y}`;
  for (let i=1;i<c.length;i++){
    const p=c[i-1],q=c[i],n=c[i+1];
    if (!n){d+=` L ${q.x} ${q.y}`;}
    else {
      const dx1=q.x-p.x,dy1=q.y-p.y,dx2=n.x-q.x,dy2=n.y-q.y;
      const l1=Math.hypot(dx1,dy1)||1,l2=Math.hypot(dx2,dy2)||1;
      const rr=Math.min(r,l1/2,l2/2);
      d+=` L ${q.x-(dx1/l1)*rr} ${q.y-(dy1/l1)*rr}`;
      d+=` Q ${q.x} ${q.y} ${q.x+(dx2/l2)*rr} ${q.y+(dy2/l2)*rr}`;
    }
  }
  return d;
};

// ── 세그먼트 목록 ───────────────────────────────────────────
const getSegments = pts => {
  const segs=[];
  for (let i=0;i<pts.length-1;i++){
    const a=pts[i],b=pts[i+1];
    segs.push({
      x1:a.x,y1:a.y,x2:b.x,y2:b.y,
      isHoriz:Math.abs(a.y-b.y)<2,
      mx:(a.x+b.x)/2,my:(a.y+b.y)/2,
      len:Math.hypot(b.x-a.x,b.y-a.y),
    });
  }
  return segs;
};

// ── PipeEdge ─────────────────────────────────────────────────
const PipeEdge = ({
  id,
  sourceX,sourceY,sourcePosition,
  targetX,targetY,targetPosition,
  source,target,
  data,selected,
}) => {
  const {getNodes, getEdges}=useReactFlow();

  const lt        = data?.lineType||"Piping";
  const ls        = LINE_STYLE[lt]||LINE_STYLE.Piping;
  const baseColor = data?.fluidSub?getFluidColor(data.fluidSub):ls.color;
  const stroke    = selected?"#f59e0b":baseColor;
  const sw        = ls.sw||2;
  const mkId      = `arrow_${id}`;

  const icStatusColor={"OPEN":"#CA8A04","IN PROGRESS":"#2563EB","CLOSED":"#16A34A","OVERDUE":"#DC2626"};
  const fluidLabel  = data?.fluidSub||"";
  const sizeLabel   = data?.sizeNum?`${data.sizeNum}A`:(data?.size||"");
  const pipingLabel = [fluidLabel,sizeLabel].filter(Boolean).join("-");
  const isSpecial   = lt==="Process Gas"||lt==="Material";
  const showLabel   = isSpecial?(data?.lineText||lt):pipingLabel;
  const icNo        = data?.ic_no||"";
  const icStatus    = data?.ic_status||"";
  // v10: IC_STATUS_FLOW 색상 우선, 없으면 구버전 색상
  const icColor     = (IC_STATUS_FLOW[icStatus]?.color) || icStatusColor[icStatus] || "#64748B";
  const ifType      = data?.ifType||"";       // v10 Interface Type
  const icdNo       = data?.icd_no||"";        // v10 ICD 번호
  const icdStatus   = data?.icd_status||"";    // v10 ICD 상태

  // 장애물 수집 (v10.8: 무관한 Area도 장애물로 포함)
  //  - source/target 노드 제외
  //  - source/target이 "속해 있는" Area(조상 포함)는 통과 허용
  //  - 그 외 Area는 관통 금지 → 라인이 무관 Area를 피해 라우팅됨
  const allNodesSnapshot = getNodes();
  const obstacles = useMemo(()=>{
    const getBox = (n) => {
      const x = n.position?.x||0, y = n.position?.y||0;
      const w = n.measured?.width  ?? n.width  ?? n.style?.width  ?? 120;
      const h = n.measured?.height ?? n.height ?? n.style?.height ?? 60;
      return { x, y, w, h };
    };
    const srcNode = allNodesSnapshot.find(n=>n.id===source);
    const tgtNode = allNodesSnapshot.find(n=>n.id===target);
    // 노드 중심이 Area bbox 안에 있는지 (소속 판정)
    const centerInside = (node, areaBox) => {
      if (!node) return false;
      const nb = getBox(node);
      const cx = nb.x + nb.w/2, cy = nb.y + nb.h/2;
      return cx >= areaBox.x && cx <= areaBox.x + areaBox.w
          && cy >= areaBox.y && cy <= areaBox.y + areaBox.h;
    };
    return allNodesSnapshot
      .filter(n=>{
        if (n.id===source || n.id===target) return false;
        if (n.type==="area") {
          const ab = getBox(n);
          // start/end가 속한 Area(중첩 조상 포함)는 장애물에서 제외
          if (centerInside(srcNode, ab) || centerInside(tgtNode, ab)) return false;
          return true; // 무관한 Area → 관통 금지
        }
        return true;
      })
      .map(n=>{
        const b = getBox(n);
        return {
          x1: b.x - PADDING,
          y1: b.y - PADDING,
          x2: b.x + b.w + PADDING,
          y2: b.y + b.h + PADDING,
        };
      });
  },[source, target,
    // 모든 비-source/target 노드 위치를 의존성에 포함 → 노드 이동 시 재계산
    allNodesSnapshot.map(n => `${n.id}:${n.position?.x},${n.position?.y}:${n.measured?.width||n.width||0}x${n.measured?.height||n.height||0}`).join("|"),
  ]);

  const storedWp = data?.waypoints||[];
  const isDragging = data?._dragging===true;

  // 경로 계산 (v10.7 결정론적 직교 라우터):
  //  - 수동 waypoints 있음 → routePipe 정규화 (사용자 조정 우선)
  //  - 그 외 → routeOrthogonal (사선 없는 L/Z 라우팅 + 박스 우회)
  //  → 모든 세그먼트 100% 수평/수직 보장, 산 모양 방지
  const pts = useMemo(()=>{
    const sPos = sourcePosition||"right";
    const tPos = targetPosition||"left";

    if (storedWp.length > 0) {
      // 사용자 조정 waypoint 사용 (정규화만 — 직각 스냅 포함)
      return routePipe(sourceX, sourceY, sPos, targetX, targetY, tPos, storedWp);
    }

    // 자동 라우팅: 결정론적 직교 (드래그 중엔 장애물 회피 생략으로 빠르게)
    const route = routeOrthogonal(
      sourceX, sourceY, sPos, targetX, targetY, tPos,
      isDragging ? [] : obstacles
    );
    // 평행 offset (같은 노드 쌍 다중 라인 분리)
    const offset = applyParallelOffset(route, id, getEdges(), source, target);
    // 중간 세그먼트에 조정용 waypoint 삽입 (선택 시 핸들로 잡기 가능)
    return isDragging ? offset : insertMidWaypoints(offset);
  },[id,source,target,sourceX,sourceY,sourcePosition,
     targetX,targetY,targetPosition,storedWp,obstacles,
     isDragging,getEdges]);

  const path = buildElbowPath(pts);
  const segs = getSegments(pts);

  const totalLen    = segs.reduce((s,g)=>s+g.len,0);

  // ── 화살표 절대 크기 (userSpaceOnUse 기준 px) ─────────
  // 라인 길이가 짧으면 작게, 길면 크게.
  // 단 라인 굵기(sw)도 반영하여 굵은 라인에서도 화살표가 보이게 보장.
  const baseLen = totalLen < 40 ? 6
                : totalLen < 80 ? 8
                : totalLen < 200 ? 10
                : 12;
  // 굵은 라인일수록 최소 크기 보장 (sw 2 → 8, sw 5 → 14)
  const arrowLen = Math.max(baseLen, sw * 2.5 + 1);
  const arrowW   = arrowLen * 0.6;
  // 화살촉 base를 line 끝점에 정확히 정렬
  const refXEnd  = arrowLen - 0.5;
  const longest     = segs.reduce((a,b)=>b.len>a.len?b:a,segs[0]||{mx:0,my:0,len:0});

  const toSVG=(svg,ev)=>{
    const pt=svg.createSVGPoint();
    pt.x=ev.clientX;pt.y=ev.clientY;
    return pt.matrixTransform(svg.getScreenCTM().inverse());
  };

  // 세그먼트 드래그: 수평→y 이동, 수직→x 이동
  const onSegDrag=useCallback((e,segIdx)=>{
    e.stopPropagation();
    const svg=e.target.closest("svg");if(!svg)return;
    const seg=segs[segIdx];
    const origPos=toSVG(svg,e);
    // v10.8: 항상 "현재 렌더된 경로(pts)" 기준으로 드래그
    // storedWp 기준이면 normalize/rectify로 점 개수가 달라져
    // 세그먼트↔waypoint 인덱스가 어긋나 사선이 생기던 버그 수정
    const initWps=pts.slice(1,-1).map(p=>({...p}));

    const onMove=mv=>{
      const cur=toSVG(svg,mv);
      const nw=initWps.map(p=>({...p}));
      const wi1=segIdx-1,wi2=segIdx;
      if(seg.isHoriz){
        // GRID(10px) 스냅 — 정돈된 좌표 유지
        const dy=snapGrid(cur.y-origPos.y);
        if(wi1>=0&&wi1<nw.length) nw[wi1]={...nw[wi1],y:initWps[wi1].y+dy};
        if(wi2>=0&&wi2<nw.length) nw[wi2]={...nw[wi2],y:(initWps[wi2]?.y??cur.y)+dy};
      } else {
        const dx=snapGrid(cur.x-origPos.x);
        if(wi1>=0&&wi1<nw.length) nw[wi1]={...nw[wi1],x:initWps[wi1].x+dx};
        if(wi2>=0&&wi2<nw.length) nw[wi2]={...nw[wi2],x:(initWps[wi2]?.x??cur.x)+dx};
      }
      const allPts=normalizePath([{x:sourceX,y:sourceY},...nw,{x:targetX,y:targetY}]);
      window.dispatchEvent(new CustomEvent("mbse:updatewaypoint",
        {detail:{id,waypoints:allPts.slice(1,-1)}}));
    };
    const onUp=()=>{window.removeEventListener("mousemove",onMove);window.removeEventListener("mouseup",onUp);};
    window.addEventListener("mousemove",onMove);
    window.addEventListener("mouseup",onUp);
  },[id,segs,storedWp,pts,sourceX,sourceY,targetX,targetY]);

  return (
    <g className="mbse-pipe-edge">
      <defs>
        {/* 화살표: userSpaceOnUse → 라인 굵기와 무관한 절대 픽셀 크기 */}
        <marker id={mkId}
          markerWidth={arrowLen} markerHeight={arrowW}
          refX={refXEnd} refY={arrowW/2}
          orient="auto" markerUnits="userSpaceOnUse">
          <polygon points={`0,0 ${arrowLen},${arrowW/2} 0,${arrowW}`}
            fill={stroke} stroke="none"/>
        </marker>
        {ls.bidir&&(
          <marker id={`${mkId}_start`}
            markerWidth={arrowLen} markerHeight={arrowW}
            refX={0.5} refY={arrowW/2}
            orient="auto-start-reverse" markerUnits="userSpaceOnUse">
            <polygon points={`0,0 ${arrowLen},${arrowW/2} 0,${arrowW}`}
              fill={stroke} stroke="none"/>
          </marker>
        )}
      </defs>

      {/* 히트 영역 — 클릭: 선택 / 더블클릭: 수동 조정 리셋(자동 라우팅 복귀) */}
      <path d={path} fill="none" stroke="transparent" strokeWidth={16}
        style={{cursor:"pointer"}}
        onDoubleClick={(e)=>{
          e.stopPropagation();
          if (storedWp.length > 0) {
            window.dispatchEvent(new CustomEvent("mbse:updatewaypoint",
              {detail:{id, waypoints: []}}));
          }
        }}>
        <title>{storedWp.length>0?"더블클릭: 자동 라우팅으로 복귀":""}</title>
      </path>

      {/* 선택 시 글로우 underlay — 어떤 라인이 선택됐는지 명확히 */}
      {selected && (
        <path d={path} fill="none"
          stroke="#f59e0b" strokeWidth={sw+6} opacity={0.25}
          strokeLinecap="round" style={{pointerEvents:"none"}}/>
      )}

      {/* 실제 라인 */}
      <path d={path} fill="none"
        stroke={stroke} strokeWidth={sw}
        strokeDasharray={ls.dash==="none"?"":ls.dash}
        markerEnd={`url(#${mkId})`}
        markerStart={ls.bidir?`url(#${mkId}_start)`:undefined}
        style={{pointerEvents:"none"}}/>

      {/* 선택 시 세그먼트 핸들 — 중앙 파란 사각형 (stub 제외) */}
      {selected&&segs.map((seg,i)=>{
        if(seg.len<16) return null;
        const isStub=i===0||i===segs.length-1;
        if(isStub) return null;
        const cursor=seg.isHoriz?"ns-resize":"ew-resize";
        return(
          <rect key={i}
            x={seg.mx-5} y={seg.my-5} width={10} height={10} rx={1}
            fill="#2563EB" stroke="white" strokeWidth={1.5}
            style={{cursor,pointerEvents:"all"}}
            onMouseDown={e=>onSegDrag(e,i)}/>
        );
      })}

      {/* 라벨 + IC/IF Type/ICD 배지 */}
      {(showLabel||icNo||ifType||icdNo)&&longest.len>30&&(
        <EdgeLabelRenderer>
          <div style={{
            position:"absolute",
            transform:`translate(-50%,-100%) translate(${longest.mx}px,${longest.my}px)`,
            display:"flex",flexDirection:"column",alignItems:"center",gap:2,
            pointerEvents:"none",
          }}>
            {/* IF Type 배지 */}
            {ifType&&IF_TYPES[ifType]&&(
              <div style={{
                fontSize:8,fontWeight:800,
                background:IF_TYPES[ifType].color,color:"#fff",
                padding:"0 5px",borderRadius:3,whiteSpace:"nowrap",letterSpacing:0.3,
              }}>{ifType}</div>
            )}
            {showLabel&&(
              <div style={{
                fontSize:10,fontWeight:isSpecial?700:500,fontFamily:"monospace",
                background:"rgba(255,255,255,0.95)",padding:"1px 6px",
                borderRadius:4,border:`1px solid ${baseColor}`,color:baseColor,
                whiteSpace:"nowrap",boxShadow:"0 1px 3px rgba(0,0,0,0.1)",
              }}>{showLabel}</div>
            )}
            {/* ICD 번호 배지 */}
            {icdNo&&(
              <div style={{
                fontSize:8,fontWeight:700,fontFamily:"monospace",
                background:"#1e293b",color:"#fff",
                padding:"0 5px",borderRadius:3,whiteSpace:"nowrap",
              }}>{icdNo}{icdStatus?` [${icdStatus}]`:""}</div>
            )}
            {icNo&&(
              <div style={{
                fontSize:9,fontWeight:700,
                background:icStatus?icColor:"#64748B",
                color:"#fff",padding:"0 5px",borderRadius:3,whiteSpace:"nowrap",
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
// ═══════════════════════════════════════════════════════════════
// IT EDGE — Area↔Area Interface Tie-in (v10.3)
// 양방향 굵은 보라색 dashed 라인 + IT-N 라벨
// 클릭 시 Definition + Check List 풀화면 모달 열림
// ═══════════════════════════════════════════════════════════════
const ITEdge = ({
  id, sourceX, sourceY, sourcePosition,
  targetX, targetY, targetPosition,
  source, target, data, selected,
}) => {
  const { getNodes, getEdges } = useReactFlow();
  const itNo = data?.itNo || "IT-?";
  const mkId = `it_arrow_${id}`;

  // v10.5: Zoom out 시 IT 라벨 자동 확대 (최대 3배 — 가장 중요한 정보라 더 크게)
  const labelScale = useZoomScale(3.0);

  // ── 장애물 수집 (IT는 Area도 회피, 자기 source/target + 조상 Area 제외) ──
  // 일반 노드(Equipment, Instrument)도 회피하여 깔끔한 여백 라우팅 보장
  const allNodes = getNodes();
  const obstacles = useMemo(() => {
    const getBox = (n) => {
      const x = n.position?.x||0, y = n.position?.y||0;
      const w = n.measured?.width  ?? n.width  ?? n.style?.width  ?? 120;
      const h = n.measured?.height ?? n.height ?? n.style?.height ?? 60;
      return { x, y, w, h };
    };
    const srcNode = allNodes.find(n=>n.id===source);
    const tgtNode = allNodes.find(n=>n.id===target);
    const centerInside = (node, areaBox) => {
      if (!node) return false;
      const nb = getBox(node);
      const cx = nb.x + nb.w/2, cy = nb.y + nb.h/2;
      return cx >= areaBox.x && cx <= areaBox.x + areaBox.w
          && cy >= areaBox.y && cy <= areaBox.y + areaBox.h;
    };
    return allNodes
      .filter(n => {
        if (n.id === source || n.id === target) return false;
        // v10.8: source/target Area를 감싸는 조상 Area는 제외
        // (Plant 안의 두 Package 연결 시 Plant가 장애물이 되어 거대 우회하는 문제 방지)
        if (n.type === "area") {
          const ab = getBox(n);
          if (centerInside(srcNode, ab) || centerInside(tgtNode, ab)) return false;
        }
        return true;
      })
      .map(n => {
        const b = getBox(n);
        // IT 라인은 Area Block 경계와 더 멀리 — 큰 PADDING (16px)
        const pad = 16;
        return {
          x1: b.x - pad,
          y1: b.y - pad,
          x2: b.x + b.w + pad,
          y2: b.y + b.h + pad,
        };
      });
  }, [source, target,
    allNodes.map(n => `${n.id}:${n.position?.x},${n.position?.y}:${n.measured?.width||n.width||0}x${n.measured?.height||n.height||0}`).join("|"),
  ]);

  const storedWp   = data?.waypoints || [];
  const isDragging = data?._dragging === true;

  // v10.5: 끝점에 수직 stub 강제 — 화살표가 핸들에 너무 붙지 않도록
  //  - source/target 포트 방향으로 최소 IT_STUB 만큼 직선 확보
  //  - 라인 시작·끝의 가독성 + 화살표 머리 여유 공간 보장
  const IT_STUB = 36;   // 일반 라인(MIN_STUB=20)보다 큼 — IT 가독성 우선
  const enforceStubs = (rawPts) => {
    if (!rawPts || rawPts.length < 2) return rawPts;
    const dirVec = (pos) => pos==="right" ? {dx:1,dy:0}
                          : pos==="left"  ? {dx:-1,dy:0}
                          : pos==="bottom"? {dx:0,dy:1}
                          : {dx:0,dy:-1};
    const sv = dirVec(sourcePosition || "top");
    const tv = dirVec(targetPosition || "top");
    const result = rawPts.map(p => ({...p}));

    // ── Source stub 확보 ──
    const sx = result[0].x, sy = result[0].y;
    const sNext = result[1];
    // 다음 점이 source 핸들 방향으로 IT_STUB 만큼 가지 못한 경우, stub 점 삽입
    const sExtX = sx + sv.dx * IT_STUB;
    const sExtY = sy + sv.dy * IT_STUB;
    // 방향 일치 확인: 다음 점이 핸들 방향으로 진행 중이면 거리 검사
    const sProgress = (sNext.x - sx) * sv.dx + (sNext.y - sy) * sv.dy;
    if (sProgress < IT_STUB - 0.5) {
      // stub 짧음 → 강제 stub 점을 두 번째 위치에 삽입
      result.splice(1, 0, { x: sExtX, y: sExtY });
    }

    // ── Target stub 확보 ──
    const tx = result[result.length-1].x, ty = result[result.length-1].y;
    const tPrev = result[result.length-2];
    const tExtX = tx + tv.dx * IT_STUB;
    const tExtY = ty + tv.dy * IT_STUB;
    const tProgress = (tPrev.x - tx) * tv.dx + (tPrev.y - ty) * tv.dy;
    if (tProgress < IT_STUB - 0.5) {
      result.splice(result.length-1, 0, { x: tExtX, y: tExtY });
    }

    return normalizePath(result);
  };

  // ── 경로 계산 (v10.7 결정론적 직교 라우터) ────────────────
  //  · 수동 waypoints → routePipe 정규화 (사용자 조정 우선)
  //  · 그 외 → routeOrthogonal + IT 전용 긴 stub
  //  → 사선 없는 L/Z 라우팅, 산 모양 방지
  const pts = useMemo(() => {
    const sPos = sourcePosition || "top";
    const tPos = targetPosition || "top";

    if (storedWp.length > 0) {
      const raw = routePipe(sourceX, sourceY, sPos, targetX, targetY, tPos, storedWp);
      return enforceStubs(raw);
    }

    // 결정론적 직교 라우팅 (드래그 중엔 장애물 회피 생략)
    const route = routeOrthogonal(
      sourceX, sourceY, sPos, targetX, targetY, tPos,
      isDragging ? [] : obstacles
    );
    // IT 전용 긴 stub 적용 + 중간 waypoint
    const stubbed = enforceStubs(route);
    return isDragging ? stubbed : insertMidWaypoints(stubbed);
  }, [id, source, target, sourceX, sourceY, sourcePosition,
      targetX, targetY, targetPosition, storedWp, obstacles, isDragging]);

  const dPath = buildElbowPath(pts);
  const segs  = getSegments(pts);

  // ── 라벨 위치: 가장 긴 세그먼트 중앙 ─────────────────────
  const longest = segs.reduce((a,b)=>b.len>a.len?b:a, segs[0]||{mx:0,my:0,len:0});

  // ── 체크리스트 진척 상태 색상 (라벨 테두리) ──────────────
  const cl = data?.checklist || {};
  const total = Object.keys(cl).length || 5;
  const done  = Object.values(cl).filter(itCatHasContent).length; // 신·구 형식 호환
  const pct   = total > 0 ? done/total : 0;
  const statusColor = pct === 1 ? "#16a34a"
                    : pct > 0    ? "#f59e0b"
                    : "#94a3b8";

  // ── stroke ──
  const stroke = selected ? "#f59e0b" : "#7c3aed";
  const sw     = selected ? 3.5 : 2;

  // ── SVG 좌표 변환 (마우스 좌표 → SVG 좌표) ──
  const toSVG = (svg, ev) => {
    const p = svg.createSVGPoint();
    p.x = ev.clientX; p.y = ev.clientY;
    return p.matrixTransform(svg.getScreenCTM().inverse());
  };

  // ── 세그먼트 드래그: 사용자 라우트 조정 ─────────────────
  // 수평 세그먼트 → 위/아래로 드래그하여 Y 이동
  // 수직 세그먼트 → 좌/우로 드래그하여 X 이동
  const onSegDrag = useCallback((e, segIdx) => {
    e.stopPropagation();
    const svg = e.target.closest("svg");
    if (!svg) return;
    const seg = segs[segIdx];
    const origPos = toSVG(svg, e);
    // v10.8: 항상 렌더된 경로(pts) 기준 — 인덱스 어긋남으로 인한 사선 방지
    const initWps = pts.slice(1, -1).map(p => ({...p}));

    const onMove = mv => {
      const cur = toSVG(svg, mv);
      const nw = initWps.map(p => ({...p}));
      const wi1 = segIdx - 1, wi2 = segIdx;
      if (seg.isHoriz) {
        // GRID(10px) 스냅 — 정돈된 좌표 유지
        const dy = snapGrid(cur.y - origPos.y);
        if (wi1 >= 0 && wi1 < nw.length) nw[wi1] = {...nw[wi1], y: initWps[wi1].y + dy};
        if (wi2 >= 0 && wi2 < nw.length) nw[wi2] = {...nw[wi2], y: (initWps[wi2]?.y ?? cur.y) + dy};
      } else {
        const dx = snapGrid(cur.x - origPos.x);
        if (wi1 >= 0 && wi1 < nw.length) nw[wi1] = {...nw[wi1], x: initWps[wi1].x + dx};
        if (wi2 >= 0 && wi2 < nw.length) nw[wi2] = {...nw[wi2], x: (initWps[wi2]?.x ?? cur.x) + dx};
      }
      const allPts = normalizePath([
        {x:sourceX, y:sourceY}, ...nw, {x:targetX, y:targetY}
      ]);
      window.dispatchEvent(new CustomEvent("mbse:updatewaypoint",
        { detail: { id, waypoints: allPts.slice(1, -1) } }));
    };
    const onUp = () => {
      window.removeEventListener("mousemove", onMove);
      window.removeEventListener("mouseup", onUp);
    };
    window.addEventListener("mousemove", onMove);
    window.addEventListener("mouseup", onUp);
  }, [id, segs, storedWp, pts, sourceX, sourceY, targetX, targetY]);

  return (
    <g className="mbse-it-edge" data-it-id={id} data-selected={selected ? "1" : "0"}>
      <defs>
        <marker id={mkId} markerWidth={9} markerHeight={7}
          refX={8} refY={3.5} orient="auto" markerUnits="userSpaceOnUse">
          <polygon points="0,0 9,3.5 0,7" fill={stroke}/>
        </marker>
        <marker id={`${mkId}_start`} markerWidth={9} markerHeight={7}
          refX={1} refY={3.5} orient="auto-start-reverse" markerUnits="userSpaceOnUse">
          <polygon points="0,0 9,3.5 0,7" fill={stroke}/>
        </marker>
      </defs>

      {/* 히트 영역 — 클릭: 선택(라우트 조정) / 더블클릭: 자동 라우팅 복귀 */}
      <path d={dPath} fill="none" stroke="transparent" strokeWidth={20}
        style={{cursor:"pointer"}}
        onDoubleClick={(e)=>{
          e.stopPropagation();
          if (storedWp.length > 0) {
            window.dispatchEvent(new CustomEvent("mbse:updatewaypoint",
              { detail: { id, waypoints: [] } }));
          }
        }}>
        <title>{storedWp.length>0?"더블클릭: 자동 라우팅으로 복귀":""}</title>
      </path>

      {/* 선택 시 글로우 underlay */}
      {selected && (
        <path d={dPath} fill="none"
          stroke="#f59e0b" strokeWidth={sw+6} opacity={0.25}
          strokeLinecap="round" style={{pointerEvents:"none"}}/>
      )}

      {/* 실제 라인 */}
      <path className="mbse-it-edge-path" d={dPath} fill="none"
        stroke={stroke} strokeWidth={sw}
        strokeDasharray="6 4"
        markerEnd={`url(#${mkId})`}
        markerStart={`url(#${mkId}_start)`}
        style={{pointerEvents:"none"}}/>

      {/* ── 세그먼트 중앙 드래그 핸들 (선택 시만, stub 제외) ──
          IT 라인 본체 클릭 → selected → 핸들 표시
          핸들을 드래그하여 해당 세그먼트의 위치 조정 가능 */}
      {selected && segs.map((seg, i) => {
        if (seg.len < 16) return null;
        const isStub = i === 0 || i === segs.length - 1;
        if (isStub) return null;
        const cursor = seg.isHoriz ? "ns-resize" : "ew-resize";
        return (
          <g key={i} className="mbse-it-segment-handle">
            {/* 외부 글로우 (시각적 강조) */}
            <circle
              cx={seg.mx} cy={seg.my} r={11}
              fill="#7c3aed" opacity={0.15}
              style={{ pointerEvents:"none" }}
            />
            {/* 메인 핸들 — 드래그로 라우트 조정 */}
            <circle
              cx={seg.mx} cy={seg.my} r={7}
              fill="#fff" stroke="#7c3aed" strokeWidth={3}
              style={{ cursor, pointerEvents:"all",
                       filter:"drop-shadow(0 1px 2px rgba(0,0,0,0.2))" }}
              onMouseDown={ev => onSegDrag(ev, i)}
            />
            {/* 중앙 점 (방향 인디케이터) */}
            <circle
              cx={seg.mx} cy={seg.my} r={3}
              fill="#7c3aed" style={{ pointerEvents:"none" }}
            />
            {/* 방향 화살표 (가로면 위아래, 세로면 좌우) */}
            {seg.isHoriz ? (
              <>
                <path d={`M ${seg.mx-4} ${seg.my-9} L ${seg.mx} ${seg.my-12} L ${seg.mx+4} ${seg.my-9}`}
                  fill="none" stroke="#7c3aed" strokeWidth={1.5} strokeLinecap="round"
                  style={{ pointerEvents:"none" }}/>
                <path d={`M ${seg.mx-4} ${seg.my+9} L ${seg.mx} ${seg.my+12} L ${seg.mx+4} ${seg.my+9}`}
                  fill="none" stroke="#7c3aed" strokeWidth={1.5} strokeLinecap="round"
                  style={{ pointerEvents:"none" }}/>
              </>
            ) : (
              <>
                <path d={`M ${seg.mx-9} ${seg.my-4} L ${seg.mx-12} ${seg.my} L ${seg.mx-9} ${seg.my+4}`}
                  fill="none" stroke="#7c3aed" strokeWidth={1.5} strokeLinecap="round"
                  style={{ pointerEvents:"none" }}/>
                <path d={`M ${seg.mx+9} ${seg.my-4} L ${seg.mx+12} ${seg.my} L ${seg.mx+9} ${seg.my+4}`}
                  fill="none" stroke="#7c3aed" strokeWidth={1.5} strokeLinecap="round"
                  style={{ pointerEvents:"none" }}/>
              </>
            )}
          </g>
        );
      })}

      {/* IT-N 라벨 (클릭 시 모달 열림) */}
      <EdgeLabelRenderer>
        <div
          className="mbse-it-edge-label"
          style={{
            position:"absolute",
            // v10.5: 줌 아웃 시 자동 확대 (최대 3배). Interface 모드 CSS 의 1.15x 와는 별개로 곱해짐
            transform:`translate(-50%,-50%) translate(${longest.mx}px,${longest.my}px) scale(${labelScale})`,
            transformOrigin:"center center",
            transition: "transform 0.15s ease-out",
            background:"#7c3aed",
            color:"#fff",
            padding:"2px 8px",
            borderRadius:10,
            fontSize:10.5,
            fontWeight:700,
            fontFamily:"monospace",
            border:`2px solid ${statusColor}`,
            boxShadow:"0 2px 5px rgba(124,58,237,0.3)",
            cursor:"pointer",
            pointerEvents:"all",
            whiteSpace:"nowrap",
            display:"flex",
            alignItems:"center",
            gap:4,
          }}
          onClick={(e) => {
            e.stopPropagation();
            window.dispatchEvent(new CustomEvent("mbse:open-it-modal", { detail:{ id } }));
          }}
          title={`${itNo} — 체크리스트 ${done}/${total} 작성됨\n• 라벨 클릭: Definition + Check List 모달 열기\n• 라인 본체 클릭: 선택 → 라우트 조정 핸들 표시`}
        >
          🔗 {itNo}
        </div>
      </EdgeLabelRenderer>
    </g>
  );
};

// ─────────────────────────────────────────────────────────────
const nodeTypes = { area:AreaNode, equipment:EquipmentNode, instrument:InstrumentNode, brench:BrenchNode };
const edgeTypes = { pipe:PipeEdge, itEdge:ITEdge };

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
        {/* v10.3: IT Line — Area↔Area Interface Tie-in */}
        <div draggable onDragStart={e=>onDragStart(e,"itLine","IT")}
          style={{ ...iS("#7c3aed"), borderTop:"1px dashed #c4b5fd", marginTop:6, paddingTop:8 }} {...hov}>
          <span style={{ display:"inline-flex", alignItems:"center", gap:3, marginRight:6 }}>
            <span style={{ color:"#7c3aed", fontSize:9 }}>◀</span>
            <span style={{ display:"inline-block",width:20,height:2.5,background:"#7c3aed",borderRadius:1,borderTop:"1px dashed transparent" }}/>
            <span style={{ color:"#7c3aed", fontSize:9 }}>▶</span>
          </span>
          🔗 IT Line (Area↔Area)
        </div>
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
  const [reqAssignee,setReqAssignee]=useState(""), [reqReview,setReqReview]=useState("");
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

  // v10.1: Scope of Supply 담당자 업데이트
  const upSos = (group, role, v) => {
    const sos = d.sos || { posco:{}, supplier:{} };
    const updated = {
      ...sos,
      [group]: { ...(sos[group]||{}), [role]: v },
    };
    onUpdateNode(sel.id, { ...d, sos: updated });
  };
  // SoS 입력 폼 렌더 (Area/Equipment 공통)
  const renderSoS = () => (
    <>
      <div style={{ fontWeight:700,fontSize:11,color:"#1f4e79",margin:"10px 0 6px",borderTop:"2px solid #ddebf7",paddingTop:7 }}>
        👷 POSCO 담당 (복수 입력 가능)
      </div>
      {SOS_POSCO_ROLES.map(r=>(
        <div key={r} style={{ marginBottom:5 }}>
          <label style={L}>{r}</label>
          <input style={{ ...I, marginBottom:0 }}
            value={d.sos?.posco?.[r]||""}
            onChange={e=>upSos("posco",r,e.target.value)}
            placeholder="홍길동, 김철수"/>
        </div>
      ))}
      <div style={{ fontWeight:700,fontSize:11,color:"#843c0c",margin:"10px 0 6px",borderTop:"2px solid #fce4d6",paddingTop:7 }}>
        🏭 Supplier 담당
      </div>
      {SOS_SUPPLIER_ROLES.map(r=>(
        <div key={r} style={{ marginBottom:5 }}>
          <label style={L}>{r}</label>
          <input style={{ ...I, marginBottom:0 }}
            value={d.sos?.supplier?.[r]||""}
            onChange={e=>upSos("supplier",r,e.target.value)}
            placeholder={`${r} 담당자`}/>
        </div>
      ))}
    </>
  );
  const addReq=()=>{
    if(!reqText.trim()) return;
    const reqs=[...(d.requirements||[]),{
      id:Date.now(), text:reqText, who:reqWho, date:reqDate,
      assignee:reqAssignee, review:reqReview,
    }];
    onUpdateNode(sel.id,{...d,requirements:reqs});
    setReqText(""); setReqWho(""); setReqAssignee(""); setReqReview("");
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
          {sel.type==="area"&&<TB name="scope" label="Scope"/>}
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
                {(() => {
                  // 표준 spec 필드
                  const standardFields = getSpecFields(d.equipType);
                  const stdKeyMap = new Map(standardFields.map(f => [f.key, f]));
                  // 시스템 예약 키 (Inspector에 별도 표시 안 함)
                  const reserved = new Set([
                    "label","itemNo","equipType","instrCategory","instrType",
                    "handles","requirements","summary","autoInlets","autoOutlets",
                    "sos","poscoStaff","engStaff","wbsCode","teamId","discipline",
                    "scopeText","inclusions","exclusions","vendorName","vendorCountry",
                    "vendorContract","vendorICA","areaType","ports","_dragging",
                    "ic_no","ic_status","ic_priority","ic_due","ic_closed",
                    "ic_resp_from","ic_resp_to","ic_remark","icd_no","ifType",
                    "icd_status","tq_no","tq_status","openItems","icd_ifaDate","icd_ifcDate",
                  ]);
                  // 데이터에 실제로 있는 키들 — 객체 입력 순서(사양서 import 순서) 유지
                  const dataKeys = Object.keys(d||{}).filter(k =>
                    !reserved.has(k) && d[k] != null && d[k] !== ""
                  );

                  // 1) 사양서/데이터에 있는 키를 먼저 배치 (등장 순서대로)
                  //    - 표준 필드면 표준 label/unit 사용
                  //    - 표준에 없으면 imported로 표시
                  const seen = new Set();
                  const result = [];
                  dataKeys.forEach(k => {
                    if (stdKeyMap.has(k)) {
                      result.push(stdKeyMap.get(k));
                    } else {
                      result.push({ key:k, label:k, unit:"", _imported:true });
                    }
                    seen.add(k);
                  });
                  // 2) 아직 값이 없는 표준 필드도 뒤에 배치 (입력 가능하도록)
                  standardFields.forEach(f => {
                    if (!seen.has(f.key)) result.push(f);
                  });
                  return result;
                })().map(({key,label,unit,_imported})=>(
                  <div key={key} style={{ marginBottom:5 }}>
                    <label style={{ ...L, marginBottom:1, display:"flex", alignItems:"center", gap:4 }}>
                      <span>{label}</span>
                      {unit && (
                        <span style={{ color:"#94a3b8",fontWeight:400,fontSize:10 }}>
                          {unit}
                        </span>
                      )}
                      {_imported && (
                        <span style={{
                          fontSize:9,background:"#fef3c7",color:"#92400e",
                          borderRadius:3,padding:"0 4px",marginLeft:"auto",fontWeight:600
                        }}>imported</span>
                      )}
                    </label>
                    <input
                      style={{ ...I, marginBottom:0 }}
                      value={(() => {
                        const v = d[key] || "";
                        if (!v) return "";
                        // 값 끝에서 단위처럼 보이는 모든 형태를 제거
                        //  - 라벨에 정의된 unit (exact match)
                        //  - 일반적인 공학 단위 (대소문자 무시, 변형 허용)
                        const commonUnits = [
                          "m³/h","m3/h","Nm³/h","Nm3/h","L/h","l/h","t/h","kg/h",
                          "m³","m3","Nm³","Nm3","kW","MW","W","kPa","MPa","Pa",
                          "Bar g","Bar","bar","barg","Bar(g)",
                          "℃","°C","C","°F","F","K",
                          "mm","cm","m","inch","\"",
                          "%","ppm","ppb",
                          "kg/cm²","kg/cm2","kgf/cm²","kgf/cm2",
                        ];
                        let stripped = String(v).trim();
                        // 1) 라벨 unit 우선 제거
                        if (unit) {
                          const re = new RegExp(`\\s*${unit.replace(/[.*+?^${}()|[\]\\]/g,'\\$&')}\\s*$`, "i");
                          stripped = stripped.replace(re, "").trim();
                        }
                        // 2) 일반 단위 제거 (긴 것부터 매칭)
                        const sorted = [...commonUnits].sort((a,b) => b.length - a.length);
                        for (const u of sorted) {
                          const re = new RegExp(`\\s*${u.replace(/[.*+?^${}()|[\]\\]/g,'\\$&')}\\s*$`, "i");
                          const next = stripped.replace(re, "").trim();
                          if (next !== stripped) { stripped = next; break; }
                        }
                        return stripped;
                      })()}
                      onChange={e => upN(key, e.target.value)}
                      onBlur={e => {
                        // 입력 완료 시점에 값에서 단위 제거하여 데이터 정리
                        const v = e.target.value;
                        if (!v) return;
                        const commonUnits = [
                          "m³/h","m3/h","Nm³/h","Nm3/h","L/h","l/h","t/h","kg/h",
                          "m³","m3","Nm³","Nm3","kW","MW","W","kPa","MPa","Pa",
                          "Bar g","Bar","bar","barg","Bar(g)",
                          "℃","°C","C","°F","F","K",
                          "mm","cm","m","inch","\"",
                          "%","ppm","ppb",
                          "kg/cm²","kg/cm2","kgf/cm²","kgf/cm2",
                        ];
                        let cleaned = String(v).trim();
                        if (unit) {
                          const re = new RegExp(`\\s*${unit.replace(/[.*+?^${}()|[\]\\]/g,'\\$&')}\\s*$`, "i");
                          cleaned = cleaned.replace(re, "").trim();
                        }
                        const sorted = [...commonUnits].sort((a,b) => b.length - a.length);
                        for (const u of sorted) {
                          const re = new RegExp(`\\s*${u.replace(/[.*+?^${}()|[\]\\]/g,'\\$&')}\\s*$`, "i");
                          const next = cleaned.replace(re, "").trim();
                          if (next !== cleaned) { cleaned = next; break; }
                        }
                        if (cleaned !== v) upN(key, cleaned);
                      }}
                      placeholder=""
                    />
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

                {/* ═══ Scope of Supply 담당자 (Excel 연동) ═══ */}
                <div style={{ marginTop:10,padding:"6px 8px",background:"#f8fafc",borderRadius:6,fontSize:10,color:"#64748b",lineHeight:1.5 }}>
                  📋 아래 담당자는 <b>Scope of Supply</b> 엑셀 시트에 출력되며, 엑셀에서 수정 후 Import하면 반영됩니다.
                </div>
                {renderSoS()}
              </>
            )}
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

                {/* ═══ v10: Interface 관리 (Type/ICD/TQ) ═══ */}
                <div style={{ fontWeight:700,fontSize:11,color:"#1d4ed8",margin:"10px 0 5px",borderTop:"2px solid #eff6ff",paddingTop:7 }}>
                  🔗 Interface 관리
                </div>

                {/* Interface Type */}
                <label style={L}>Interface Type</label>
                <select style={S} value={d.ifType||""} onChange={e=>upE("ifType",e.target.value)}>
                  <option value="">Select Type</option>
                  {Object.entries(IF_TYPES).map(([k,v])=>(
                    <option key={k} value={k}>{v.label}</option>
                  ))}
                </select>
                {d.ifType&&IF_TYPES[d.ifType]&&(
                  <div style={{ display:"flex",alignItems:"center",gap:6,marginBottom:7,padding:"4px 8px",background:"#f8fafc",borderRadius:4,borderLeft:`3px solid ${IF_TYPES[d.ifType].color}` }}>
                    <span style={{ fontSize:10,color:"#475569" }}>{IF_TYPES[d.ifType].desc}</span>
                    <span style={{ fontSize:9,marginLeft:"auto",fontWeight:700,color:IF_TYPES[d.ifType].color }}>위험 {IF_TYPES[d.ifType].risk}</span>
                  </div>
                )}

                {/* IC Status (편집 가능) */}
                <label style={L}>IC Status</label>
                <select style={S} value={d.ic_status||""} onChange={e=>upE("ic_status",e.target.value)}>
                  <option value="">Select Status</option>
                  {IC_STATUS_LIST.map(s=><option key={s} value={s}>{s}</option>)}
                </select>
                {d.ic_status&&IC_STATUS_FLOW[d.ic_status]&&(
                  <div style={{ marginBottom:7,padding:"4px 8px",borderRadius:4,background:"#f8fafc",borderLeft:`3px solid ${IC_STATUS_FLOW[d.ic_status].color}` }}>
                    <div style={{ fontSize:10,color:"#475569" }}>{IC_STATUS_FLOW[d.ic_status].desc}</div>
                    <div style={{ fontSize:10,fontWeight:700,color:IC_STATUS_FLOW[d.ic_status].color,marginTop:1 }}>→ {IC_STATUS_FLOW[d.ic_status].next}</div>
                  </div>
                )}

                {/* ICD 관리 */}
                <label style={L}>ICD No.</label>
                <input style={I} value={d.icd_no||""} onChange={e=>upE("icd_no",e.target.value)} placeholder="e.g. ICD-FBR-ESF-001"/>
                <label style={L}>ICD Status</label>
                <select style={S} value={d.icd_status||""} onChange={e=>upE("icd_status",e.target.value)}>
                  <option value="">Select</option>
                  {ICD_STATUS_LIST.map(s=><option key={s} value={s}>{s}</option>)}
                </select>
                <div style={{ display:"flex",gap:5 }}>
                  <div style={{ flex:1 }}>
                    <label style={L}>IFA Date</label>
                    <input type="date" style={I} value={d.icd_ifaDate||""} onChange={e=>upE("icd_ifaDate",e.target.value)}/>
                  </div>
                  <div style={{ flex:1 }}>
                    <label style={L}>IFC Date</label>
                    <input type="date" style={I} value={d.icd_ifcDate||""} onChange={e=>upE("icd_ifcDate",e.target.value)}/>
                  </div>
                </div>

                {/* TQ 추적 */}
                <label style={L}>TQ No.</label>
                <input style={I} value={d.tq_no||""} onChange={e=>upE("tq_no",e.target.value)} placeholder="e.g. TQ-001"/>
                <label style={L}>TQ Status</label>
                <select style={S} value={d.tq_status||""} onChange={e=>upE("tq_status",e.target.value)}>
                  <option value="">Select</option>
                  {TQ_STATUS_LIST.map(s=><option key={s} value={s}>{s}</option>)}
                </select>
                <label style={L}>Open Items (미결사항)</label>
                <textarea style={{ ...I,height:40,resize:"vertical" }} value={d.openItems||""} onChange={e=>upE("openItems",e.target.value)} placeholder="미결 Action Item..."/>

                {/* IC Register 연동 정보 (Import 후 표시) */}
                {(d.ic_no||d.ic_priority||d.ic_due) && (
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
                {(r.assignee || r.review) && (
                  <div style={{ marginTop:4,paddingTop:4,borderTop:"1px dashed #fde68a" }}>
                    {r.assignee && (
                      <div style={{ fontSize:10,color:"#1e40af",marginBottom:1 }}>
                        <span style={{ fontWeight:600 }}>담당자:</span> {r.assignee}
                      </div>
                    )}
                    {r.review && (
                      <div style={{ fontSize:10,color:"#166534",lineHeight:1.4 }}>
                        <span style={{ fontWeight:600 }}>검토결과:</span> {r.review}
                      </div>
                    )}
                  </div>
                )}
                <button onClick={()=>{ const rs=(d.requirements||[]).filter(x=>x.id!==r.id); onUpdateNode(sel.id,{...d,requirements:rs}); }}
                  style={{ background:"none",border:"none",color:"#dc2626",cursor:"pointer",fontSize:10,padding:"2px 0",marginTop:2 }}>× remove</button>
              </div>
            ))}
            {!(d.requirements||[]).length&&<div style={{ color:"#94a3b8",fontSize:11,marginBottom:10 }}>등록된 요구사항 없음</div>}
            <div style={{ borderTop:"1px solid #f1f5f9",paddingTop:8 }}>
              <textarea style={{ ...I,height:56,resize:"vertical" }} placeholder="Requirement..." value={reqText} onChange={e=>setReqText(e.target.value)}/>
              <input style={I} placeholder="Stakeholder" value={reqWho} onChange={e=>setReqWho(e.target.value)}/>
              <input type="date" style={I} value={reqDate} onChange={e=>setReqDate(e.target.value)}/>
              <input style={I} placeholder="담당자" value={reqAssignee} onChange={e=>setReqAssignee(e.target.value)}/>
              <textarea style={{ ...I,height:40,resize:"vertical" }} placeholder="검토결과" value={reqReview} onChange={e=>setReqReview(e.target.value)}/>
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

        {/* ═══ v10: Scope 탭 (Area 전용) ═══ */}
        {tab==="scope"&&isNode&&sel.type==="area"&&(
          <>
            <div style={{ fontWeight:700,fontSize:11,color:"#1d4ed8",marginBottom:6 }}>📐 Scope 정의</div>
            <label style={L}>WBS Code</label>
            <input style={I} value={d.wbsCode||""} onChange={e=>upN("wbsCode",e.target.value)} placeholder="e.g. FBR-UT-WT-001"/>

            <div style={{ display:"flex",gap:5 }}>
              <div style={{ flex:1 }}>
                <label style={L}>Team</label>
                <select style={S} value={d.teamId||""} onChange={e=>upN("teamId",e.target.value)}>
                  <option value="">Select</option>
                  <option value="FBR">FBR</option>
                  <option value="ESF">ESF</option>
                  <option value="Common">Common</option>
                </select>
              </div>
              <div style={{ flex:1 }}>
                <label style={L}>Discipline</label>
                <select style={S} value={d.discipline||""} onChange={e=>upN("discipline",e.target.value)}>
                  <option value="">Select</option>
                  {VENDOR_DISCIPLINES.map(x=><option key={x}>{x}</option>)}
                </select>
              </div>
            </div>

            <label style={L}>범위 기술 (Scope)</label>
            <textarea style={{ ...I,height:50,resize:"vertical" }} value={d.scopeText||""} onChange={e=>upN("scopeText",e.target.value)} placeholder="범위 기술..."/>
            <label style={L}>포함 범위 (Inclusions)</label>
            <textarea style={{ ...I,height:36,resize:"vertical" }} value={d.inclusions||""} onChange={e=>upN("inclusions",e.target.value)}/>
            <label style={L}>제외 범위 (Exclusions)</label>
            <textarea style={{ ...I,height:36,resize:"vertical" }} value={d.exclusions||""} onChange={e=>upN("exclusions",e.target.value)}/>

            {/* ── Vendor 정보 ── */}
            <div style={{ fontWeight:700,fontSize:11,color:"#1d4ed8",margin:"10px 0 6px",borderTop:"2px solid #eff6ff",paddingTop:7 }}>🏭 Vendor 정보</div>
            <label style={L}>회사명</label>
            <input style={I} value={d.vendorName||""} onChange={e=>upN("vendorName",e.target.value)} placeholder="e.g. Primetals"/>
            <div style={{ display:"flex",gap:5 }}>
              <div style={{ flex:1 }}>
                <label style={L}>국가</label>
                <input style={I} value={d.vendorCountry||""} onChange={e=>upN("vendorCountry",e.target.value)} placeholder="Austria"/>
              </div>
              <div style={{ flex:1 }}>
                <label style={L}>계약번호</label>
                <input style={I} value={d.vendorContract||""} onChange={e=>upN("vendorContract",e.target.value)}/>
              </div>
            </div>
            <label style={L}>Interface Coordinator</label>
            <input style={I} value={d.vendorICA||""} onChange={e=>upN("vendorICA",e.target.value)} placeholder="ICA 담당자"/>

            {/* ── POSCO 담당 (복수) ── */}
            <div style={{ fontWeight:700,fontSize:11,color:"#1d4ed8",margin:"10px 0 6px",borderTop:"2px solid #eff6ff",paddingTop:7 }}>👷 POSCO 담당</div>
            {POSCO_ORGS.map(org=>(
              <div key={org} style={{ marginBottom:6 }}>
                <label style={L}>{org} 담당자 (쉼표로 복수 입력)</label>
                <input style={I}
                  value={(d.poscoStaff?.[org])||""}
                  onChange={e=>upN("poscoStaff",{...(d.poscoStaff||{}),[org]:e.target.value})}
                  placeholder="홍길동, 김철수"/>
              </div>
            ))}

            {/* ── Engineering 담당 (복수) ── */}
            <div style={{ fontWeight:700,fontSize:11,color:"#1d4ed8",margin:"10px 0 6px",borderTop:"2px solid #eff6ff",paddingTop:7 }}>🔧 Engineering 담당</div>
            {VENDOR_DISCIPLINES.map(disc=>(
              <div key={disc} style={{ marginBottom:6 }}>
                <label style={L}>{disc} 담당자 (쉼표로 복수)</label>
                <input style={I}
                  value={(d.engStaff?.[disc])||""}
                  onChange={e=>upN("engStaff",{...(d.engStaff||{}),[disc]:e.target.value})}
                  placeholder="담당자 복수 입력"/>
              </div>
            ))}

            {/* ═══ Scope of Supply 담당자 (Excel 연동) ═══ */}
            <div style={{ marginTop:10,padding:"6px 8px",background:"#f8fafc",borderRadius:6,fontSize:10,color:"#64748b",lineHeight:1.5 }}>
              📋 아래 담당자는 <b>Scope of Supply</b> 엑셀 시트에 출력되며, 엑셀에서 수정 후 Import하면 반영됩니다.
            </div>
            {renderSoS()}
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

// ═══════════════════════════════════════════════════════════════
// IT LINE MODAL — Definition + Check List (v10.3)
// 캡쳐의 양식 디지털화: 좌측 Definition, 우측 Check List
// ═══════════════════════════════════════════════════════════════
const ITLineModal = ({ edge, nodes, allITEdges, onSave, onClose, onNavigate, onDelete }) => {
  const d = edge?.data || {};

  // 양쪽 Area 노드 찾기
  const sourceNode = nodes.find(n => n.id === edge?.source);
  const targetNode = nodes.find(n => n.id === edge?.target);
  const sourceName = sourceNode?.data?.label || sourceNode?.id || "?";
  const targetName = targetNode?.data?.label || targetNode?.id || "?";

  // 로컬 편집 상태
  const [itNoSuffix, setItNoSuffix] = useState(() =>
    (d.itNo || "IT-1").replace(/^IT-/, "")
  );
  const [description, setDescription] = useState(d.description || "");
  const [supplyByArea, setSupplyByArea] = useState(() => d.supplyByArea || [
    { areaName: sourceName, supplier: "P-ENC" },
    { areaName: targetName, supplier: "P-ENC" },
  ]);
  const [functionalDesc, setFunctionalDesc] = useState(d.functionalDesc || "");
  // v10.9: 영향인자 = {text, status, owner} 객체 (구버전 문자열 자동 변환)
  const [influenceFactors, setInfluenceFactors] = useState(() =>
    (d.influenceFactors && d.influenceFactors.length > 0
      ? d.influenceFactors : [""]
    ).map(normalizeITFactor)
  );
  const [checklist, setChecklist] = useState(() => {
    // v10.10: 모든 카테고리를 신형식으로 정규화 (구버전 content → items 자동 마이그레이션)
    const obj = {};
    IT_CHECKLIST_CATEGORIES.forEach(c => {
      obj[c.key] = normalizeITChecklistCat(d.checklist?.[c.key]);
    });
    return obj;
  });
  // 삭제 확인 다이얼로그
  const [showDeleteConfirm, setShowDeleteConfirm] = useState(false);

  // 저장
  const handleSave = () => {
    const itNo = `IT-${itNoSuffix.trim() || "1"}`;
    // 빈 항목 제거 후 저장
    const cleanChecklist = {};
    Object.entries(checklist).forEach(([k, c]) => {
      cleanChecklist[k] = { ...c, items: (c.items||[]).filter(it => it.text.trim()) };
    });
    onSave(edge.id, {
      itNo,
      description,
      supplyByArea,
      functionalDesc,
      influenceFactors: influenceFactors.filter(x => x.text.trim()),
      checklist: cleanChecklist,
    });
  };

  // 영향인자 추가/삭제/필드수정
  const updateFactor = (i, field, v) => {
    const next = [...influenceFactors];
    next[i] = { ...next[i], [field]: v };
    setInfluenceFactors(next);
  };
  const addFactor = () => setInfluenceFactors([...influenceFactors, { text:"", status:"확인 중", owner:"" }]);
  const removeFactor = (i) => setInfluenceFactors(influenceFactors.filter((_, idx) => idx !== i));

  // 공급 구분별 변경
  const updateSupply = (i, field, v) => {
    const next = [...supplyByArea];
    next[i] = { ...next[i], [field]: v };
    setSupplyByArea(next);
  };

  // 체크리스트 변경 (카테고리 필드: supplier/supplierOwner/responsible/responsibleOwner)
  const updateChecklist = (catKey, field, v) => {
    setChecklist({ ...checklist, [catKey]: { ...checklist[catKey], [field]: v } });
  };
  // v10.10: 체크리스트 개별 항목 조작
  const updateChecklistItem = (catKey, i, field, v) => {
    const items = [...(checklist[catKey]?.items || [])];
    items[i] = { ...items[i], [field]: v };
    setChecklist({ ...checklist, [catKey]: { ...checklist[catKey], items } });
  };
  const addChecklistItem = (catKey) => {
    const items = [...(checklist[catKey]?.items || []), { text:"", status:"확인 중" }];
    setChecklist({ ...checklist, [catKey]: { ...checklist[catKey], items } });
  };
  const removeChecklistItem = (catKey, i) => {
    const items = (checklist[catKey]?.items || []).filter((_, idx) => idx !== i);
    setChecklist({ ...checklist, [catKey]: { ...checklist[catKey], items } });
  };

  // IT 라인 네비게이션
  const currentIdx = allITEdges.findIndex(e => e.id === edge.id);
  const prevEdge = currentIdx > 0 ? allITEdges[currentIdx - 1] : null;
  const nextEdge = currentIdx < allITEdges.length - 1 ? allITEdges[currentIdx + 1] : null;

  const IT_DESC = `${sourceName} ↔ ${targetName}`;

  // 공통 스타일
  const sectionH = { fontSize:12, fontWeight:700, color:"#1e293b", marginBottom:6,
                     paddingBottom:4, borderBottom:"2px solid #e2e8f0" };
  const fieldL   = { fontSize:11, fontWeight:600, color:"#475569", marginBottom:3, display:"block" };
  const inputS   = { width:"100%", padding:"5px 8px", border:"1px solid #cbd5e1",
                     borderRadius:4, fontSize:11, boxSizing:"border-box", fontFamily:"inherit" };

  return (
    <div onClick={onClose} style={{
      position:"fixed", inset:0, background:"rgba(15,23,42,0.6)", zIndex:1000,
      display:"flex", alignItems:"center", justifyContent:"center", padding:20,
    }}>
      <div onClick={e=>e.stopPropagation()} style={{
        width:"min(1380px, 96vw)", height:"min(90vh, 880px)",
        background:"#fff", borderRadius:10, display:"flex", flexDirection:"column",
        boxShadow:"0 20px 50px rgba(0,0,0,0.3)", overflow:"hidden",
      }}>
        {/* ── HEADER ── */}
        <div style={{
          padding:"12px 20px", background:"#7c3aed", color:"#fff",
          display:"flex", alignItems:"center", gap:14,
        }}>
          <span style={{ fontSize:18, fontWeight:800 }}>🔗 Interface Tie-in 관리</span>
          <div style={{ display:"flex", alignItems:"center", gap:6, background:"rgba(255,255,255,0.2)", padding:"3px 10px", borderRadius:6 }}>
            <span style={{ fontSize:13, fontWeight:700 }}>IT-</span>
            <input value={itNoSuffix} onChange={e=>setItNoSuffix(e.target.value)}
              style={{ width:60, padding:"2px 6px", borderRadius:4, border:"none",
                       fontSize:13, fontWeight:700, color:"#1e293b", textAlign:"center" }}/>
          </div>
          <span style={{ fontSize:13, opacity:0.9 }}>{IT_DESC}</span>
          <div style={{ marginLeft:"auto", display:"flex", gap:6 }}>
            {prevEdge && (
              <button onClick={()=>onNavigate(prevEdge.id)} style={{
                background:"rgba(255,255,255,0.2)", border:"none", color:"#fff",
                padding:"5px 12px", borderRadius:5, cursor:"pointer", fontSize:11,
              }}>◀ 이전</button>
            )}
            <span style={{ fontSize:11, opacity:0.85, alignSelf:"center" }}>
              {currentIdx + 1} / {allITEdges.length}
            </span>
            {nextEdge && (
              <button onClick={()=>onNavigate(nextEdge.id)} style={{
                background:"rgba(255,255,255,0.2)", border:"none", color:"#fff",
                padding:"5px 12px", borderRadius:5, cursor:"pointer", fontSize:11,
              }}>다음 ▶</button>
            )}
            {/* 삭제 버튼 — 확인 후 IT 라인 삭제 */}
            <button onClick={()=>setShowDeleteConfirm(true)} style={{
              background:"#dc2626", border:"none", color:"#fff",
              padding:"5px 12px", borderRadius:5, cursor:"pointer", fontSize:11, fontWeight:700,
              marginLeft:10,
            }} title="이 IT Line 삭제">🗑 삭제</button>
            {/* 닫기 — 변경사항 무시하고 닫기 */}
            <button onClick={onClose} style={{
              background:"rgba(255,255,255,0.25)", border:"none", color:"#fff",
              padding:"5px 14px", borderRadius:5, cursor:"pointer", fontSize:11, fontWeight:700,
            }}>✕ 닫기</button>
          </div>
        </div>

        {/* ── BODY (2-column) ── */}
        <div style={{ flex:1, display:"grid", gridTemplateColumns:"1fr 1.3fr", gap:0, overflow:"hidden" }}>
          {/* ─── LEFT: Definition ─── */}
          <div style={{ padding:"16px 20px", overflowY:"auto", borderRight:"1px solid #e2e8f0", background:"#fafbfd" }}>
            <div style={{ fontSize:14, fontWeight:800, color:"#7c3aed", marginBottom:12,
                          padding:"6px 10px", background:"#f3e8ff", borderRadius:6,
                          display:"flex", alignItems:"center", gap:6 }}>
              📐 Definition
            </div>

            <div style={{ marginBottom:14 }}>
              <label style={fieldL}>IC&TI Description</label>
              <div style={{ padding:"8px 10px", background:"#fff", border:"1px solid #cbd5e1",
                            borderRadius:5, fontSize:12, fontWeight:600, color:"#1e293b" }}>
                <span style={{ color:"#2563eb" }}>{sourceName}</span>
                <span style={{ margin:"0 8px", color:"#7c3aed" }}>↔</span>
                <span style={{ color:"#dc2626" }}>{targetName}</span>
              </div>
              <input value={description} onChange={e=>setDescription(e.target.value)}
                style={{ ...inputS, marginTop:5 }}
                placeholder="추가 설명 (선택)"/>
            </div>

            <div style={{ marginBottom:14 }}>
              <label style={fieldL}>공급 구분별 Inter-connection</label>
              {supplyByArea.map((sup, i) => (
                <div key={i} style={{ display:"flex", gap:6, marginBottom:5, alignItems:"center" }}>
                  <input value={sup.areaName} onChange={e=>updateSupply(i, "areaName", e.target.value)}
                    style={{ ...inputS, flex:1.5 }} placeholder="Area"/>
                  <span style={{ color:"#94a3b8", fontSize:11 }}>:</span>
                  <input value={sup.supplier} onChange={e=>updateSupply(i, "supplier", e.target.value)}
                    style={{ ...inputS, flex:1 }} placeholder="공급처 (e.g. P-ENC)"/>
                </div>
              ))}
            </div>

            <div style={{ marginBottom:14 }}>
              <label style={fieldL}>기능 운영간 Inter-connection (서술)</label>
              <textarea value={functionalDesc} onChange={e=>setFunctionalDesc(e.target.value)}
                style={{ ...inputS, minHeight:80, resize:"vertical", lineHeight:1.4 }}
                placeholder="예: HyREX TUF & Bucket Elevator 는 Diverter Chute & Oxide Bin 의 가동조건에 따름..."/>
            </div>

            <div>
              <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", marginBottom:5 }}>
                <label style={{ ...fieldL, marginBottom:0 }}>Process 간 주요 영향인자</label>
                <button onClick={addFactor} style={{
                  background:"#7c3aed", color:"#fff", border:"none", borderRadius:4,
                  padding:"2px 8px", fontSize:10, cursor:"pointer", fontWeight:700,
                }}>+ 추가</button>
              </div>
              {influenceFactors.map((f, i) => {
                const st = IT_FACTOR_STATUS.find(s => s.value === f.status) || IT_FACTOR_STATUS[0];
                return (
                  <div key={i} style={{ display:"flex", gap:5, marginBottom:4, alignItems:"center" }}>
                    <span style={{ fontSize:11, color:"#7c3aed", fontWeight:700, width:18 }}>{i+1}.</span>
                    <input value={f.text} onChange={e=>updateFactor(i, "text", e.target.value)}
                      style={{ ...inputS, flex:2.2 }} placeholder="영향인자 (예: HyREX TUF 가동 여부 확인)"/>
                    {/* 진행현황 — 상태별 색상 표시 */}
                    <select value={f.status} onChange={e=>updateFactor(i, "status", e.target.value)}
                      style={{
                        ...inputS, flex:"0 0 76px", fontWeight:700,
                        color: st.color, borderColor: st.color+"80",
                      }}>
                      {IT_FACTOR_STATUS.map(s=><option key={s.value} value={s.value}>{s.value}</option>)}
                    </select>
                    {/* 담당자 — 추진반/ENC/DX 조직별 1~2명 키인 */}
                    <input value={f.owner} onChange={e=>updateFactor(i, "owner", e.target.value)}
                      style={{ ...inputS, flex:1.2 }} placeholder="담당자 (예: 추진반 홍길동)"/>
                    <button onClick={()=>removeFactor(i)} style={{
                      background:"none", border:"none", color:"#dc2626", cursor:"pointer",
                      fontSize:14, padding:"0 4px",
                    }}>×</button>
                  </div>
                );
              })}
            </div>
          </div>

          {/* ─── RIGHT: Check List ─── */}
          <div style={{ padding:"16px 20px", overflowY:"auto", background:"#fff" }}>
            <div style={{ fontSize:14, fontWeight:800, color:"#facc15", marginBottom:12,
                          padding:"6px 10px", background:"#fef9c3", borderRadius:6,
                          color:"#854d0e", display:"flex", alignItems:"center", gap:6 }}>
              ✓ Check List
            </div>

            {/* IT-x, Description, 영향인자 요약 */}
            <div style={{ marginBottom:12, padding:"8px 10px", background:"#f8fafc",
                          borderRadius:5, fontSize:11, color:"#475569",
                          border:"1px solid #e2e8f0" }}>
              <div><b style={{ color:"#7c3aed" }}>Interface No.</b> IT-{itNoSuffix}</div>
              <div><b style={{ color:"#7c3aed" }}>Description.</b> {IT_DESC}</div>
              <div style={{ marginTop:4 }}>
                <b style={{ color:"#7c3aed" }}>주요 영향인자:</b>{" "}
                {influenceFactors.filter(f=>f.text.trim()).length === 0
                  ? <i style={{color:"#94a3b8"}}>입력 필요</i>
                  : influenceFactors.filter(f=>f.text.trim()).map((f, i) => {
                      const st = IT_FACTOR_STATUS.find(s=>s.value===f.status) || IT_FACTOR_STATUS[0];
                      return (
                        <span key={i} style={{ display:"inline-flex", alignItems:"center", gap:3, marginRight:8 }}>
                          {f.text}
                          <span style={{
                            fontSize:9, fontWeight:700, color:"#fff", background:st.color,
                            borderRadius:3, padding:"0 4px",
                          }}>{f.status}</span>
                          {f.owner && <span style={{ fontSize:10, color:"#64748b" }}>({f.owner})</span>}
                        </span>
                      );
                    })}
              </div>
            </div>

            {/* 각 카테고리 */}
            {IT_CHECKLIST_CATEGORIES.map(cat => {
              const cl = checklist[cat.key] || normalizeITChecklistCat(null);
              const items = cl.items || [];
              const filled = items.some(it => (it.text||"").trim());
              return (
                <div key={cat.key} style={{
                  marginBottom:10, border:`1px solid ${cat.hdrFc}40`, borderRadius:6, overflow:"hidden",
                }}>
                  {/* 헤더 1줄: 카테고리명 + 작성 상태 + 항목 추가 */}
                  <div style={{
                    padding:"6px 10px", background:cat.hdrBg, color:cat.hdrFc,
                    display:"flex", alignItems:"center", gap:8,
                    fontSize:11, fontWeight:700,
                  }}>
                    <span style={{ flex:1 }}>{cat.label}</span>
                    <span style={{
                      fontSize:9, padding:"1px 6px", borderRadius:3,
                      background: filled ? "#16a34a" : "#cbd5e1", color:"#fff", fontWeight:700,
                    }}>{filled ? `작성됨 ${items.filter(it=>it.text.trim()).length}건` : "미작성"}</span>
                    <button onClick={()=>addChecklistItem(cat.key)} style={{
                      background:cat.hdrFc, color:"#fff", border:"none", borderRadius:3,
                      padding:"1px 7px", fontSize:9, cursor:"pointer", fontWeight:700,
                    }}>+ 항목</button>
                  </div>
                  {/* 헤더 2줄: 공급 select + 담당자 / 주관 select + 담당자 */}
                  <div style={{
                    padding:"5px 10px", background:cat.hdrBg+"80",
                    display:"grid", gridTemplateColumns:"auto 1fr auto 1fr",
                    gap:6, alignItems:"center", fontSize:10,
                    borderTop:`1px dashed ${cat.hdrFc}30`,
                  }}>
                    <select value={cl.supplier} onChange={e=>updateChecklist(cat.key,"supplier",e.target.value)}
                      style={{ padding:"2px 6px", borderRadius:3, border:"1px solid #cbd5e1",
                               fontSize:10, fontWeight:600, color:"#1e293b" }}>
                      <option value="">공급</option>
                      {IT_SUPPLIER_OPTIONS.map(x=><option key={x}>{x}</option>)}
                    </select>
                    <input value={cl.supplierOwner||""} onChange={e=>updateChecklist(cat.key,"supplierOwner",e.target.value)}
                      style={{ padding:"2px 6px", borderRadius:3, border:"1px solid #cbd5e1",
                               fontSize:10, color:"#1e293b", minWidth:0 }}
                      placeholder="공급 담당자"/>
                    <select value={cl.responsible} onChange={e=>updateChecklist(cat.key,"responsible",e.target.value)}
                      style={{ padding:"2px 6px", borderRadius:3, border:"1px solid #cbd5e1",
                               fontSize:10, fontWeight:600, color:"#1e293b" }}>
                      <option value="">주관</option>
                      {IT_RESPONSIBLE_OPTIONS.map(x=><option key={x}>{x}</option>)}
                    </select>
                    <input value={cl.responsibleOwner||""} onChange={e=>updateChecklist(cat.key,"responsibleOwner",e.target.value)}
                      style={{ padding:"2px 6px", borderRadius:3, border:"1px solid #cbd5e1",
                               fontSize:10, color:"#1e293b", minWidth:0 }}
                      placeholder="주관 담당자"/>
                  </div>
                  {/* 항목 리스트: 내용 + 진행현황 + 삭제 */}
                  <div style={{ padding: items.length>0 ? "5px 10px 7px" : "0" }}>
                    {items.length === 0 && (
                      <div style={{ padding:"8px 10px", fontSize:10, color:"#94a3b8", fontStyle:"italic" }}>
                        "+ 항목" 버튼으로 점검/논의 항목을 추가하세요
                      </div>
                    )}
                    {items.map((it, i) => {
                      const st = IT_FACTOR_STATUS.find(s => s.value === it.status) || IT_FACTOR_STATUS[0];
                      return (
                        <div key={i} style={{ display:"flex", gap:5, marginTop:i===0?2:4, alignItems:"center" }}>
                          <span style={{ fontSize:10, color:cat.hdrFc, fontWeight:700, width:16 }}>{i+1}.</span>
                          <input value={it.text} onChange={e=>updateChecklistItem(cat.key, i, "text", e.target.value)}
                            style={{ flex:1, padding:"4px 7px", border:"1px solid #cbd5e1",
                                     borderRadius:4, fontSize:10.5, boxSizing:"border-box", minWidth:0 }}
                            placeholder={`${cat.label} 점검/논의 항목...`}/>
                          <select value={it.status} onChange={e=>updateChecklistItem(cat.key, i, "status", e.target.value)}
                            style={{ flex:"0 0 72px", padding:"4px 4px", borderRadius:4, fontSize:10,
                                     fontWeight:700, color:st.color, border:`1px solid ${st.color}80` }}>
                            {IT_FACTOR_STATUS.map(s=><option key={s.value} value={s.value}>{s.value}</option>)}
                          </select>
                          <button onClick={()=>removeChecklistItem(cat.key, i)} style={{
                            background:"none", border:"none", color:"#dc2626", cursor:"pointer",
                            fontSize:13, padding:"0 3px",
                          }}>×</button>
                        </div>
                      );
                    })}
                  </div>
                </div>
              );
            })}
          </div>
        </div>

        {/* ── FOOTER ── */}
        <div style={{ padding:"10px 20px", borderTop:"1px solid #e2e8f0",
                      display:"flex", justifyContent:"flex-end", gap:8, background:"#f8fafc" }}>
          <button onClick={onClose} style={{
            padding:"7px 18px", background:"#fff", border:"1px solid #cbd5e1",
            borderRadius:5, cursor:"pointer", fontSize:12, fontWeight:600, color:"#475569",
          }}>취소</button>
          <button onClick={handleSave} style={{
            padding:"7px 22px", background:"#7c3aed", border:"none", color:"#fff",
            borderRadius:5, cursor:"pointer", fontSize:12, fontWeight:700,
          }}>💾 저장</button>
        </div>

        {/* ── 삭제 확인 다이얼로그 ── */}
        {showDeleteConfirm && (
          <div style={{
            position:"absolute", inset:0, background:"rgba(15,23,42,0.55)",
            display:"flex", alignItems:"center", justifyContent:"center", zIndex:10,
          }} onClick={()=>setShowDeleteConfirm(false)}>
            <div onClick={e=>e.stopPropagation()} style={{
              width:420, background:"#fff", borderRadius:10, overflow:"hidden",
              boxShadow:"0 20px 50px rgba(0,0,0,0.4)",
            }}>
              <div style={{
                padding:"12px 18px", background:"#dc2626", color:"#fff",
                display:"flex", alignItems:"center", gap:8,
              }}>
                <span style={{ fontSize:20 }}>⚠</span>
                <span style={{ fontSize:14, fontWeight:800 }}>IT Line 삭제 확인</span>
              </div>
              <div style={{ padding:"18px 20px" }}>
                <div style={{ fontSize:13, color:"#1e293b", marginBottom:10, fontWeight:700 }}>
                  정말 삭제하시겠습니까?
                </div>
                <div style={{ fontSize:12, color:"#475569", lineHeight:1.6, marginBottom:8 }}>
                  <b style={{ color:"#7c3aed" }}>IT-{itNoSuffix}</b>
                  {" — "}{sourceName} ↔ {targetName}
                </div>
                <div style={{
                  fontSize:11, color:"#92400e", background:"#fef3c7",
                  padding:"8px 12px", borderRadius:5, lineHeight:1.5,
                  border:"1px solid #fcd34d",
                }}>
                  이 작업은 되돌릴 수 없습니다.<br/>
                  Definition · Check List 입력 내용이 모두 사라집니다.
                  <br/><span style={{ color:"#64748b", fontSize:10 }}>
                    (단축키 Ctrl+Z 로 잠시는 복구 가능)
                  </span>
                </div>
              </div>
              <div style={{
                padding:"10px 18px", borderTop:"1px solid #e2e8f0", background:"#f8fafc",
                display:"flex", justifyContent:"flex-end", gap:8,
              }}>
                <button onClick={()=>setShowDeleteConfirm(false)} style={{
                  padding:"7px 18px", background:"#fff", border:"1px solid #cbd5e1",
                  borderRadius:5, cursor:"pointer", fontSize:12, fontWeight:600, color:"#475569",
                }}>취소</button>
                <button onClick={()=>{
                  setShowDeleteConfirm(false);
                  if (onDelete) onDelete(edge.id);
                }} style={{
                  padding:"7px 18px", background:"#dc2626", border:"none", color:"#fff",
                  borderRadius:5, cursor:"pointer", fontSize:12, fontWeight:700,
                }}>🗑 삭제 진행</button>
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
};

// ═══════════════════════════════════════════════════════════════
// SEARCH MODAL — 노드/엣지 전역 검색 + 카메라 이동
// 단축키 Ctrl+F 또는 툴바 🔍 버튼으로 호출
// ═══════════════════════════════════════════════════════════════
const SearchModal = ({ nodes, edges, onSelect, onClose }) => {
  const [query, setQuery] = useState("");
  const [activeIdx, setActiveIdx] = useState(0);
  const inputRef = useRef(null);
  const listRef = useRef(null);

  useEffect(() => {
    // 모달 열면 자동 포커스
    inputRef.current?.focus();
  }, []);

  // ── 검색 인덱스 구축 (메모이즈) ─────────────────────────
  // 각 항목은 { type, id, label, sub, fields[] } 구조
  // fields[] 에 매칭 대상 텍스트가 모두 들어가서 검색 가능
  const index = useMemo(() => {
    const items = [];
    nodes.forEach(n => {
      const d = n.data || {};
      if (n.type === "area") {
        items.push({
          type: "area", id: n.id,
          label: d.label || "(이름 없음)",
          sub:   `${d.areaType||"Area"}${d.vendorName?` · ${d.vendorName}`:""}`,
          icon:  d.areaType==="Plant" ? "🏭"
               : d.areaType==="System"? "⚙️"
               : d.areaType==="Package"? "📦" : "📐",
          fields: [d.label, d.areaType, d.wbsCode, d.vendorName, d.scopeText, d.teamId, d.discipline].filter(Boolean),
        });
      } else if (n.type === "equipment") {
        items.push({
          type: "equipment", id: n.id,
          label: d.label || "(이름 없음)",
          sub:   `${d.itemNo||""}${d.itemNo&&d.equipType?" · ":""}${d.equipType||""}`,
          icon:  "⚙️",
          fields: [d.label, d.itemNo, d.equipType, d.material, d.capacity, d.summary].filter(Boolean),
        });
      } else if (n.type === "instrument") {
        items.push({
          type: "instrument", id: n.id,
          label: d.label || "(이름 없음)",
          sub:   `${d.itemNo||""}${d.itemNo&&d.instrType?" · ":""}${d.instrType||""}`,
          icon:  "📊",
          fields: [d.label, d.itemNo, d.instrCategory, d.instrType].filter(Boolean),
        });
      } else if (n.type === "brench") {
        // Brench는 라벨이 없으니 _hint나 ID만 검색
        items.push({
          type: "brench", id: n.id,
          label: "(Branch)",
          sub: d._hint || n.id,
          icon: "🔀",
          fields: [d._hint, n.id].filter(Boolean),
        });
      }
    });
    edges.forEach(e => {
      const d = e.data || {};
      if (e.type === "itEdge") {
        const srcN = nodes.find(x => x.id === e.source);
        const tgtN = nodes.find(x => x.id === e.target);
        const desc = `${srcN?.data?.label||srcN?.id||"?"} ↔ ${tgtN?.data?.label||tgtN?.id||"?"}`;
        items.push({
          type: "itEdge", id: e.id,
          label: d.itNo || "IT-?",
          sub: desc,
          icon: "🔗",
          fields: [d.itNo, d.description, d.functionalDesc,
                   ...(d.influenceFactors||[]).map(f=>typeof f==="string"?f:(f?.text||"")),
                   desc].filter(Boolean),
        });
      } else {
        // pipe / connection
        if (!d.ic_no && !d.icd_no && !d.tq_no && !d.lineText && !d.fluidSub && !d.serialNo) return;
        const label = d.ic_no || d.icd_no || d.lineText || d.serialNo || (d.fluidSub||"Line");
        const sub = [d.fluidSub, d.sizeNum?`${d.sizeNum}A`:d.size, d.lineType, d.ic_status]
                    .filter(Boolean).join(" · ");
        items.push({
          type: "pipe", id: e.id,
          label,
          sub,
          icon: d.lineType==="Process Gas"?"💨"
              : d.lineType==="Material"?"🟫"
              : "═",
          fields: [d.ic_no, d.icd_no, d.tq_no, d.lineText, d.fluidSub, d.serialNo, d.ic_status, d.ifType, d.lineType].filter(Boolean),
        });
      }
    });
    return items;
  }, [nodes, edges]);

  // ── 검색 결과 필터링 (대소문자 무시, 부분 매칭, 점수 정렬) ──
  const results = useMemo(() => {
    const q = query.trim().toLowerCase();
    if (!q) return [];
    const matched = [];
    index.forEach(item => {
      let bestField = "";
      let bestScore = 0;
      for (const f of item.fields) {
        const fl = String(f).toLowerCase();
        if (fl === q) {
          if (10 > bestScore) { bestScore = 10; bestField = f; }
        } else if (fl.startsWith(q)) {
          if (5 > bestScore) { bestScore = 5; bestField = f; }
        } else if (fl.includes(q)) {
          if (3 > bestScore) { bestScore = 3; bestField = f; }
        }
      }
      if (bestScore > 0) matched.push({ ...item, score: bestScore, matchedField: bestField });
    });
    matched.sort((a,b) => b.score - a.score);
    return matched.slice(0, 50);
  }, [index, query]);

  // 결과가 바뀌면 active index 리셋
  useEffect(() => { setActiveIdx(0); }, [query]);

  // 키보드 네비게이션
  const handleKeyDown = (e) => {
    if (e.key === "Escape") { e.preventDefault(); onClose(); return; }
    if (e.key === "ArrowDown") {
      e.preventDefault();
      setActiveIdx(i => Math.min(i+1, results.length-1));
    } else if (e.key === "ArrowUp") {
      e.preventDefault();
      setActiveIdx(i => Math.max(i-1, 0));
    } else if (e.key === "Enter") {
      e.preventDefault();
      if (results[activeIdx]) {
        onSelect(results[activeIdx]);
      }
    }
  };

  // 활성 항목 자동 스크롤
  useEffect(() => {
    const el = listRef.current?.children?.[activeIdx];
    if (el) el.scrollIntoView({ block:"nearest" });
  }, [activeIdx]);

  // 매칭된 부분을 hi-light
  const highlight = (text) => {
    if (!query) return text;
    const t = String(text);
    const lower = t.toLowerCase();
    const idx = lower.indexOf(query.toLowerCase());
    if (idx === -1) return t;
    return (
      <>
        {t.slice(0, idx)}
        <span style={{ background:"#fde047", color:"#713f12", fontWeight:700, padding:"0 1px", borderRadius:2 }}>
          {t.slice(idx, idx+query.length)}
        </span>
        {t.slice(idx+query.length)}
      </>
    );
  };

  const typeColors = {
    area:       { bg:"#dbeafe", fc:"#1e40af" },
    equipment:  { bg:"#e0f2fe", fc:"#075985" },
    instrument: { bg:"#dcfce7", fc:"#166534" },
    brench:     { bg:"#fef3c7", fc:"#92400e" },
    pipe:       { bg:"#f3e8ff", fc:"#6b21a8" },
    itEdge:     { bg:"#ede9fe", fc:"#5b21b6" },
  };
  const typeLabel = {
    area:"Area", equipment:"Equip.", instrument:"Instr.",
    brench:"Branch", pipe:"Line", itEdge:"IT",
  };

  return (
    <div
      onClick={onClose}
      style={{
        position:"fixed", inset:0, background:"rgba(15,23,42,0.55)", zIndex:1000,
        display:"flex", alignItems:"flex-start", justifyContent:"center", paddingTop:"10vh",
      }}
    >
      <div
        onClick={e => e.stopPropagation()}
        style={{
          width:580, maxHeight:"75vh", background:"#fff", borderRadius:10,
          boxShadow:"0 20px 50px rgba(0,0,0,0.3)", overflow:"hidden",
          display:"flex", flexDirection:"column",
        }}
      >
        {/* 검색 입력 */}
        <div style={{
          padding:"14px 18px", borderBottom:"1px solid #e2e8f0",
          display:"flex", alignItems:"center", gap:10,
        }}>
          <span style={{ fontSize:18 }}>🔍</span>
          <input
            ref={inputRef}
            value={query}
            onChange={e => setQuery(e.target.value)}
            onKeyDown={handleKeyDown}
            placeholder="노드명, Item No., IC No., IT-N, 라인 번호로 검색..."
            style={{
              flex:1, fontSize:14, padding:"4px 6px",
              border:"none", outline:"none", color:"#1e293b",
            }}
          />
          <span style={{ fontSize:11, color:"#94a3b8" }}>
            {query ? `${results.length}건` : `전체 ${index.length}건`}
          </span>
          <button onClick={onClose} style={{
            background:"#f1f5f9", border:"none", color:"#475569",
            padding:"3px 9px", borderRadius:4, cursor:"pointer", fontSize:10, fontWeight:700,
          }}>ESC</button>
        </div>

        {/* 결과 목록 */}
        <div ref={listRef} style={{ flex:1, overflowY:"auto", padding:"4px 0" }}>
          {!query && (
            <div style={{ padding:30, textAlign:"center", color:"#94a3b8", fontSize:12, lineHeight:1.8 }}>
              <div style={{ fontSize:32, marginBottom:8 }}>🔍</div>
              <div>이름, Item No., IC No., IT-N, 라인 번호 등으로 검색하세요</div>
              <div style={{ fontSize:10, marginTop:10, color:"#cbd5e1" }}>
                ↑↓ 이동 · Enter 선택 · Esc 닫기
              </div>
            </div>
          )}
          {query && results.length === 0 && (
            <div style={{ padding:30, textAlign:"center", color:"#94a3b8", fontSize:12 }}>
              검색 결과 없음
            </div>
          )}
          {results.map((r, i) => {
            const tc = typeColors[r.type] || { bg:"#f1f5f9", fc:"#475569" };
            const active = i === activeIdx;
            return (
              <div
                key={`${r.type}-${r.id}`}
                onClick={() => onSelect(r)}
                onMouseEnter={() => setActiveIdx(i)}
                style={{
                  padding:"8px 18px", cursor:"pointer",
                  borderLeft: active ? "3px solid #2563eb" : "3px solid transparent",
                  background: active ? "#eff6ff" : "transparent",
                  display:"flex", alignItems:"center", gap:10,
                  transition: "background 0.1s ease",
                }}
              >
                <span style={{ fontSize:18, width:24, textAlign:"center" }}>{r.icon}</span>
                <div style={{ flex:1, minWidth:0 }}>
                  <div style={{ fontSize:13, fontWeight:700, color:"#1e293b", whiteSpace:"nowrap", overflow:"hidden", textOverflow:"ellipsis" }}>
                    {highlight(r.label)}
                  </div>
                  <div style={{ fontSize:11, color:"#64748b", whiteSpace:"nowrap", overflow:"hidden", textOverflow:"ellipsis" }}>
                    {highlight(r.sub)}
                    {r.matchedField && !String(r.sub||"").toLowerCase().includes(query.toLowerCase()) &&
                     !String(r.label||"").toLowerCase().includes(query.toLowerCase()) && (
                      <span style={{ marginLeft:6, color:"#94a3b8" }}>
                        ⤷ {highlight(String(r.matchedField).slice(0,40))}
                      </span>
                    )}
                  </div>
                </div>
                <span style={{
                  fontSize:10, padding:"2px 7px", borderRadius:3,
                  background: tc.bg, color: tc.fc, fontWeight:700, whiteSpace:"nowrap",
                }}>
                  {typeLabel[r.type]}
                </span>
              </div>
            );
          })}
        </div>

        {/* 하단 안내 */}
        {results.length > 0 && (
          <div style={{
            padding:"6px 18px", borderTop:"1px solid #e2e8f0", background:"#f8fafc",
            fontSize:10, color:"#94a3b8", display:"flex", justifyContent:"space-between",
          }}>
            <span>↑↓ 이동 · Enter 선택 · Esc 닫기</span>
            <span>{activeIdx+1} / {results.length}</span>
          </div>
        )}
      </div>
    </div>
  );
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
      //  - 패턴 매칭 후 값에서 잔여 노이즈 제거
      //  - "(m³/h) : 600" → "600"
      //  - "(barg) : 1.5" → "1.5"
      //  - "(℃) : 35"     → "35"
      //  - "600 m³/h"     → "600"
      const cleanValue = (raw) => {
        if (!raw) return "";
        let v = String(raw).trim();
        // 1) "(단위) :" 또는 "(단위):" 패턴 앞쪽에 있으면 제거
        //    "(m³/h) : 600" 또는 "(m3/h): 600"
        v = v.replace(/^\s*\([^)]*\)\s*[:：]?\s*/, "");
        // 2) 시작 부분의 ":" 만 남은 경우 제거
        v = v.replace(/^\s*[:：]\s*/, "");
        // 3) 후행 "(단위)" 제거: "600 (m³/h)"
        v = v.replace(/\s*\([^)]*\)\s*$/, "");
        // 4) 후행 일반 단위 strip
        const commonUnits = [
          "m³/h","m3/h","Nm³/h","Nm3/h","L/h","l/h","t/h","kg/h","g/l","g/L",
          "m³","m3","Nm³","Nm3","kW","MW","W","kPa","MPa","Pa",
          "Bar g","Bar","bar","barg","Bar(g)",
          "℃","°C","°F","K",
          "mm","cm","m","inch",
          "%","ppm","ppb","dB","dB(A)",
          "kg/cm²","kg/cm2","kgf/cm²","kgf/cm2",
        ];
        const sorted = [...commonUnits].sort((a,b) => b.length - a.length);
        for (const u of sorted) {
          const re = new RegExp(`\\s*${u.replace(/[.*+?^${}()|[\]\\]/g,'\\$&')}\\s*$`, "i");
          const next = v.replace(re, "").trim();
          if (next !== v) { v = next; break; }
        }
        return v.trim();
      };

      const extractVal = (patterns) => {
        for (const p of patterns) {
          const m = block.match(new RegExp(p + "\\s*[\\:：]?\\s*([^\\n\\r]+)","i"));
          if (m) {
            const cleaned = cleanValue(m[1].replace(/\s*\/\s*/g," / "));
            if (cleaned && cleaned !== "-" && cleaned !== "TBD") return cleaned;
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

const ConnModal = ({ onConfirm, onCancel, defaultType="Piping", canBeIT=false }) => {
  const [lt,setLt]=useState(defaultType);
  useEffect(()=>setLt(defaultType),[defaultType]);
  const isIT = lt === "IT Line (Area↔Area)";
  return (
    <div style={{ position:"fixed",inset:0,background:"rgba(0,0,0,0.45)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:9999 }}>
      <div style={{ background:"#fff",borderRadius:10,padding:24,minWidth:280,boxShadow:"0 8px 30px rgba(0,0,0,0.18)" }}>
        <div style={{ fontWeight:800,fontSize:15,marginBottom:14,color:"#0f172a" }}>New Connection</div>
        <label style={{ fontSize:12,color:"#64748b",display:"block",marginBottom:4 }}>Line Type</label>
        <select value={lt} onChange={e=>setLt(e.target.value)} style={{ width:"100%",padding:"6px 8px",border:"1px solid #e2e8f0",borderRadius:6,fontSize:13,marginBottom:16 }}>
          {[...CONNECTION_LIST,...CONVEYOR_LIST].map(c=><option key={c}>{c}</option>)}
          {canBeIT && <option value="IT Line (Area↔Area)">🔗 IT Line (Interface Tie-in)</option>}
        </select>
        {/* 색상 미리보기 */}
        <div style={{ display:"flex",alignItems:"center",gap:8,marginBottom:16 }}>
          {isIT ? (
            <>
              <div style={{ width:50,height:3,background:"#7c3aed",borderRadius:2,borderTop:"1px dashed transparent" }}/>
              <span style={{ fontSize:11,color:"#7c3aed",fontWeight:700 }}>양방향 IT 라인 — Area↔Area</span>
            </>
          ) : (
            <>
              <div style={{ width:50,height:LINE_STYLE[lt]?.sw||2,background:LINE_STYLE[lt]?.color||"#94a3b8",borderRadius:2 }}/>
              <span style={{ fontSize:11,color:"#64748b" }}>{lt}</span>
            </>
          )}
        </div>
        <div style={{ display:"flex",gap:8 }}>
          <button onClick={()=>onConfirm(lt)} style={{ flex:1,background: isIT?"#7c3aed":"#1d4ed8",color:"#fff",border:"none",borderRadius:6,padding:"8px",cursor:"pointer",fontWeight:700 }}>Create</button>
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
  const { screenToFlowPosition, fitView, setCenter, getZoom }=useReactFlow();
  const [nodes,setNodes,onNodesChange]=useNodesState([]);
  const [edges,setEdges,onEdgesChange]=useEdgesState([]);
  const [sel,setSel]       = useState(null);
  const [modal,setModal]   = useState(false);
  const [modalDefault,setModalDefault] = useState("Piping");
  const [canBeIT, setCanBeIT] = useState(false);
  // v10.3: IT Line modal — 클릭한 IT edge id
  const [itModalEdgeId, setItModalEdgeId] = useState(null);
  // v10.3+: Interface 모드 토글 — IT 라인 강조, 다른 라인 페이드
  const [interfaceMode, setInterfaceMode] = useState(false);
  // v10.3+: 특정 Area의 IT 목록 팝업
  const [itListAreaId, setItListAreaId] = useState(null);
  // v10.4: 전역 검색 모달 (Ctrl+F)
  const [showSearch, setShowSearch] = useState(false);
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
        if(n) setNodes(normalizeAreaZIndex(n));
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
        // 진단: requirements 보유 노드 카운트 (import 직후 변경이 사라지는지 확인용)
        const reqCount = nodes.reduce((sum,n)=>sum+((n.data?.requirements||[]).length),0);
        console.log(`[AutoSave] 노드 ${nodes.length}, 엣지 ${edges.length}, Requirements 총 ${reqCount}건 저장됨`);
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
    // v10.3: IT 라인 라벨 클릭 시 모달 열기
    const openIT = (ev) => {
      const id = ev.detail?.id;
      if (id) setItModalEdgeId(id);
    };
    window.addEventListener("mbse:open-it-modal", openIT);
    // v10.3+: Area 의 IT 카운트 배지 클릭 시 목록 팝업 열기
    const openITList = (ev) => {
      const areaId = ev.detail?.areaId;
      if (areaId) setItListAreaId(areaId);
    };
    window.addEventListener("mbse:open-it-list", openITList);
    return()=>{
      window.removeEventListener("mbse:updatelabel",fn);
      window.removeEventListener("mbse:updatewaypoint",fn);
      window.removeEventListener("mbse:open-it-modal", openIT);
      window.removeEventListener("mbse:open-it-list", openITList);
    };
  },[setNodes,setEdges]);

  // v10.3+: edges 변경 시 각 Area 노드의 itCount 자동 주입 ──────
  // (AreaNode가 직접 edges를 못 보므로 data 에 카운트를 채워줌)
  useEffect(() => {
    const itEdges = edges.filter(e => e.type === "itEdge");
    if (itEdges.length === 0) {
      // 모든 area 의 itCount 가 이미 0이면 패스
      const hasAny = nodes.some(n => n.type === "area" && (n.data?.itCount||0) > 0);
      if (!hasAny) return;
    }
    const counts = {};
    itEdges.forEach(e => {
      counts[e.source] = (counts[e.source]||0) + 1;
      counts[e.target] = (counts[e.target]||0) + 1;
    });
    setNodes(ns => ns.map(n => {
      if (n.type !== "area") return n;
      const next = counts[n.id] || 0;
      if ((n.data?.itCount||0) === next) return n;
      return { ...n, data: { ...n.data, itCount: next } };
    }));
  }, [edges]); // eslint-disable-line

  // DRAG
  const onDragStart=useCallback((e,cat,sub)=>{ e.dataTransfer.setData("mbse/cat",cat); e.dataTransfer.setData("mbse/sub",sub); e.dataTransfer.effectAllowed="move"; },[]);
  const onDragOver=useCallback(e=>{ e.preventDefault(); e.dataTransfer.dropEffect="move"; },[]);
  const onDrop=useCallback(e=>{
    e.preventDefault();
    const cat=e.dataTransfer.getData("mbse/cat"),sub=e.dataTransfer.getData("mbse/sub");
    if(!cat) return;
    const pos=screenToFlowPosition({ x:e.clientX,y:e.clientY });
    if(cat==="area"){
      // zIndex 계층: Plant(-40) → System(-30) → Package(-20) → Item(-10)
      // 큰 영역이 항상 뒤에 위치, 안의 작은 영역과 Equipment(0)가 항상 클릭 가능
      const zMap = { Plant:-40, System:-30, Package:-20, Item:-10 };
      setNodes(ns=>[...ns,{ id:uid("area"),type:"area",position:pos,style:{ width:260,height:180 },
        data:{ areaType:sub,label:"",summary:"",requirements:[],autoInlets:[],autoOutlets:[],handles:["top","bottom","left","right"] },
        zIndex: zMap[sub] ?? -10 }]);
    } else if(cat==="itLine"){
      // IT Line은 두 Area 노드를 연결해야 생성되므로 단순 드롭으로는 만들지 않음
      // 안내: Area↔Area 핸들 연결 시 자동으로 IT Line 옵션이 모달에 표시됨
      setSaveMsg("ℹ IT Line은 두 Area 사이의 핸들을 직접 끌어 연결하세요");
      setTimeout(()=>setSaveMsg(""), 3500);
      return;
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
    // ── IT 포트 식별 (핸들 ID가 "it_" prefix) ───────────────
    const srcIsIT = (params.sourceHandle||"").startsWith("it_");
    const tgtIsIT = (params.targetHandle||"").startsWith("it_");
    const srcNode = nodes.find(n=>n.id===params.source);
    const tgtNode = nodes.find(n=>n.id===params.target);
    const bothArea = srcNode?.type==="area" && tgtNode?.type==="area";

    // ── Case 1: 양쪽 모두 IT 포트 → 즉시 IT 라인 생성 ──────
    if (srcIsIT && tgtIsIT && bothArea) {
      setEdges(es => {
        const itEdges = es.filter(e => e.type === "itEdge");
        let maxNum = 0;
        itEdges.forEach(e => {
          const m = (e.data?.itNo||"").match(/^IT-(\d+)$/);
          if (m) maxNum = Math.max(maxNum, parseInt(m[1],10));
        });
        const itNo = `IT-${maxNum + 1}`;
        const newId = uid("it");
        // 다음 tick 에서 자동으로 모달 열기
        setTimeout(() => setItModalEdgeId(newId), 50);
        return addEdge({
          ...params, id:newId, type:"itEdge",
          data:{
            itNo, description:"", supplyByArea:[], functionalDesc:"",
            influenceFactors:[], checklist: makeDefaultITChecklist(),
          },
        }, es);
      });
      return;
    }

    // ── Case 2: 한쪽만 IT 포트 → 연결 거부 (혼용 방지) ─────
    if (srcIsIT !== tgtIsIT) {
      setSaveMsg("⚠ IT 포트(보라색)는 IT 포트끼리만 연결할 수 있어요");
      setTimeout(()=>setSaveMsg(""), 3500);
      return;
    }

    // ── Case 3: 일반 포트끼리 → 기존 모달 (Piping 등 선택) ──
    connRef.current=params;
    setCanBeIT(bothArea);   // Area↔Area 면 IT 옵션도 표시 (호환성)
    setModalDefault(bothArea ? "IT Line (Area↔Area)" : "Piping");
    setModal(true);
  },[nodes,setEdges]);

  const confirmConn=useCallback(lineType=>{
    const p=connRef.current; if(!p) return;
    if (lineType === "IT Line (Area↔Area)") {
      // 자동 IT-N 번호 부여 (기존 IT 라인 중 최대값+1)
      setEdges(es => {
        const itEdges = es.filter(e => e.type === "itEdge");
        let maxNum = 0;
        itEdges.forEach(e => {
          const m = (e.data?.itNo||"").match(/^IT-(\d+)$/);
          if (m) maxNum = Math.max(maxNum, parseInt(m[1],10));
        });
        const itNo = `IT-${maxNum + 1}`;
        return addEdge({
          ...p, id:uid("it"), type:"itEdge",
          data:{
            itNo, description:"", supplyByArea:[], functionalDesc:"",
            influenceFactors:[], checklist: makeDefaultITChecklist(),
          },
        }, es);
      });
    } else {
      setEdges(es=>addEdge({ ...p,id:uid("e"),type:"pipe",data:{ lineType } },es));
    }
    setModal(false); connRef.current=null;
  },[setEdges]);

  // ── IT 라인 끝점 재연결 (드래그로 이동) ──────────────────────
  // ReactFlow가 reconnectRadius 안에서 핸들을 잡으면 onReconnect 호출
  // IT 라인은 IT 포트끼리만, 같은 노드가 아니어야 한다는 제약 검증
  const reconnectingRef = useRef({ valid:false, oldEdge:null });
  const onReconnectStart = useCallback((_, edge) => {
    reconnectingRef.current = { valid:true, oldEdge:edge };
  }, []);
  const onReconnectEnd = useCallback((_, edge) => {
    // valid가 false면 새 연결이 안 됐다는 뜻 → 원본 유지 (자동 복원됨)
    reconnectingRef.current = { valid:false, oldEdge:null };
  }, []);
  const onReconnectIT = useCallback((oldEdge, newConnection) => {
    // IT 라인 전용 — 일반 엣지는 이전에 reconnect 비활성이라 그대로 동작
    if (oldEdge.type === "itEdge") {
      // 양쪽 모두 IT 포트인지 확인
      const srcIsIT = (newConnection.sourceHandle||"").startsWith("it_");
      const tgtIsIT = (newConnection.targetHandle||"").startsWith("it_");
      if (!srcIsIT || !tgtIsIT) {
        setSaveMsg("⚠ IT 포트(보라색)에만 연결할 수 있습니다");
        setTimeout(()=>setSaveMsg(""), 3000);
        return;  // 변경 안 함
      }
      const srcN = nodes.find(n=>n.id===newConnection.source);
      const tgtN = nodes.find(n=>n.id===newConnection.target);
      if (srcN?.type !== "area" || tgtN?.type !== "area") {
        setSaveMsg("⚠ IT 라인은 Area 간에만 연결 가능합니다");
        setTimeout(()=>setSaveMsg(""), 3000);
        return;
      }
      if (newConnection.source === newConnection.target) {
        setSaveMsg("⚠ 같은 Area로 자기연결 불가");
        setTimeout(()=>setSaveMsg(""), 3000);
        return;
      }
    }
    // 유효한 재연결 — source/target/handles 업데이트, 데이터는 보존
    reconnectingRef.current.valid = true;
    setEdges(es => es.map(e => e.id === oldEdge.id ? {
      ...e,
      source:       newConnection.source,
      target:       newConnection.target,
      sourceHandle: newConnection.sourceHandle,
      targetHandle: newConnection.targetHandle,
    } : e));
  }, [nodes, setEdges]);

  const onNodeClick=useCallback((_,n)=>setSel(n),[]);
  // v10.5: IT 라인 본체 클릭 → 선택 상태(라우트 조정 모드)
  //        IT 라벨(🔗 IT-N) 클릭 → 모달 (라벨 onClick 에서 dispatch)
  const onEdgeClick=useCallback((_,e)=>{
    setSel(e);
  },[]);
  const onPaneClick=useCallback(()=>setSel(null),[]);

  // ── Live re-routing: 드래그 시작 시 연결 엣지에 _dragging 플래그 ──
  const rafIdRef = useRef(null);
  const onNodeDragStart = useCallback((_, dragNode) => {
    setEdges(es => es.map(e =>
      (e.source === dragNode.id || e.target === dragNode.id)
        ? { ...e, data: { ...e.data, _dragging: true } }
        : e
    ));
  }, [setEdges]);

  // ── Smart Guide + Live re-routing (60fps RAF throttle) ──────────
  const onNodeDrag = useCallback((_evt, dragNode) => {
    // RAF로 throttle — 16ms당 1회만 처리
    if (rafIdRef.current) return;
    rafIdRef.current = requestAnimationFrame(() => {
      rafIdRef.current = null;
    });

    const others = nodes.filter(n => n.id !== dragNode.id && n.type !== "area");
    if (others.length === 0) { setGuides([]); return; }

    const da = getNodeAnchors(dragNode);
    const newGuides = [];
    let snapX = null, snapY = null;

    others.forEach(other => {
      const oa = getNodeAnchors(other);
      const canvasW = 16000, canvasH = 16000;

      const xPairs = [
        [da.left, oa.left], [da.left, oa.centerX], [da.left, oa.right],
        [da.centerX, oa.left], [da.centerX, oa.centerX], [da.centerX, oa.right],
        [da.right, oa.left], [da.right, oa.centerX], [da.right, oa.right],
      ];
      xPairs.forEach(([dv, ov]) => {
        if (Math.abs(dv - ov) <= GUIDE_TOL) {
          if (snapX === null) snapX = ov - (dv - da.left);
          newGuides.push({ x1: ov, y1: 0, x2: ov, y2: canvasH });
        }
      });

      const yPairs = [
        [da.top, oa.top], [da.top, oa.middleY], [da.top, oa.bottom],
        [da.middleY, oa.top], [da.middleY, oa.middleY], [da.middleY, oa.bottom],
        [da.bottom, oa.top], [da.bottom, oa.middleY], [da.bottom, oa.bottom],
      ];
      yPairs.forEach(([dv, ov]) => {
        if (Math.abs(dv - ov) <= GUIDE_TOL) {
          if (snapY === null) snapY = ov - (dv - da.top);
          newGuides.push({ x1: 0, y1: ov, x2: canvasW, y2: ov });
        }
      });
    });

    setGuides(newGuides);

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
      // — 키 순서를 block(=사양서) 순서대로 보존
      const specData = {};
      Object.entries(block).forEach(([k,v]) => {
        if (!["itemNos","normalizedKeys","equipName","quantity"].includes(k)) specData[k] = v;
      });
      // 설비명도 업데이트 (기존 값 없을 때만)
      if (!n.data.label && block.equipName) specData.label = block.equipName;

      // 키 순서 재구성:
      //  1) specData(=사양서 순서) 키를 먼저 배치
      //  2) 그 다음 기존 data 중 specData에 없는 키 (시스템 키, 기존 보존 데이터)
      // → Inspector에서 사양서 순서대로 표시됨
      const merged = {};
      // 1단계: spec 키 먼저
      Object.entries(specData).forEach(([k,v]) => { merged[k] = v; });
      // 2단계: 기존 data 중 spec에 없는 키만 추가
      Object.entries(n.data || {}).forEach(([k,v]) => {
        if (!(k in merged)) merged[k] = v;
      });

      return { ...n, data: merged };
    }));
    setShowSpecImport(false);
    setSaveMsg(`✅ 사양서 ${parsedBlocks.length}건 적용 완료`);
    setTimeout(()=>setSaveMsg(""),3000);
  },[setNodes]);

  const onNodeDragStop = useCallback((_, dragNode) => {
    setGuides([]);
    // 드래그 종료 → _dragging 해제 + waypoints 초기화 (A* 재계산)
    setEdges(es => es.map(e => {
      if (e.source === dragNode.id || e.target === dragNode.id) {
        const newData = { ...e.data, waypoints: [] };
        delete newData._dragging;
        return { ...e, data: newData };
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

  // v10.4: 검색 결과 선택 → 카메라 이동 + 항목 강조
  const handleSearchSelect = useCallback((item) => {
    setShowSearch(false);
    // Node 인 경우: 노드 중심으로 이동
    if (["area","equipment","instrument","brench"].includes(item.type)) {
      const node = nodes.find(n => n.id === item.id);
      if (!node) return;
      const w = node.measured?.width  ?? node.width  ?? node.style?.width  ?? 120;
      const h = node.measured?.height ?? node.height ?? node.style?.height ?? 60;
      const cx = (node.position?.x||0) + w/2;
      const cy = (node.position?.y||0) + h/2;
      // 현재 줌이 너무 작으면 0.7 이상으로
      const targetZoom = Math.max(getZoom(), 0.7);
      setCenter(cx, cy, { zoom: targetZoom, duration: 500 });
      // 노드 선택 상태 표시
      setNodes(ns => ns.map(n => ({ ...n, selected: n.id === item.id })));
      setSel(node);
    }
    // Edge 인 경우: source/target 중간점으로 이동
    else if (["pipe","itEdge"].includes(item.type)) {
      const edge = edges.find(e => e.id === item.id);
      if (!edge) return;
      const srcN = nodes.find(n => n.id === edge.source);
      const tgtN = nodes.find(n => n.id === edge.target);
      if (!srcN || !tgtN) return;
      const sw = srcN.measured?.width  ?? srcN.width  ?? 120;
      const sh = srcN.measured?.height ?? srcN.height ?? 60;
      const tw = tgtN.measured?.width  ?? tgtN.width  ?? 120;
      const th = tgtN.measured?.height ?? tgtN.height ?? 60;
      const sx = (srcN.position?.x||0) + sw/2;
      const sy = (srcN.position?.y||0) + sh/2;
      const tx = (tgtN.position?.x||0) + tw/2;
      const ty = (tgtN.position?.y||0) + th/2;
      const targetZoom = Math.max(getZoom(), 0.6);
      setCenter((sx+tx)/2, (sy+ty)/2, { zoom: targetZoom, duration: 500 });
      // 엣지 선택 상태 표시
      setEdges(es => es.map(e => ({ ...e, selected: e.id === item.id })));
      setSel(edge);
      // IT 라인이면 모달도 함께 열기
      if (item.type === "itEdge") {
        setTimeout(() => setItModalEdgeId(item.id), 550);
      }
    }
  }, [nodes, edges, setNodes, setEdges, setCenter, getZoom]);

  // 키보드 단축키
  useEffect(()=>{
    const fn=e=>{
      const tag=document.activeElement?.tagName?.toLowerCase();
      const inInput=tag==="input"||tag==="textarea"||tag==="select";
      // v10.4: Ctrl+F → 전역 검색 모달 (inInput 무시 — 어디서든 호출 가능)
      if (e.ctrlKey && e.key.toLowerCase()==="f") {
        e.preventDefault();
        setShowSearch(true);
        return;
      }
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
    r.onload=ev=>{ try{ const p=JSON.parse(ev.target.result); if(p.nodes)setNodes(normalizeAreaZIndex(p.nodes)); if(p.edges)setEdges(p.edges); }catch{ alert("Invalid JSON"); } };
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
      if (err.code === "DRM_DETECTED") {
        alert(
          "이 파일은 회사 보안 정책(DRM)에 의해 암호화되어 import할 수 없습니다.\n\n" +
          "해결 방법:\n" +
          "1. Excel에서 파일을 연 뒤 '다른 이름으로 저장' → 데스크탑 등 외부 폴더에 저장\n" +
          "2. 또는 새 엑셀 파일에 데이터를 복사·붙여넣기 후 외부 폴더에 저장한 뒤 다시 시도\n" +
          "3. 회사 IT 보안팀에 해당 파일 DRM 예외 처리 신청"
        );
        setXlsxMsg("⚠ DRM 암호화 파일 — 외부 폴더에 새로 저장 후 시도");
        setTimeout(()=>setXlsxMsg(""), 6000);
      } else {
        setXlsxMsg("오류: " + err.message);
        setTimeout(()=>setXlsxMsg(""), 4000);
      }
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

          {/* v10.4: 🔍 Search 버튼 (단축키 Ctrl+F) */}
          <button onClick={()=>setShowSearch(true)}
            style={{
              background:"#1e293b", color:"#fbbf24",
              border:"1px solid #475569", borderRadius:5,
              padding:"3px 10px", cursor:"pointer",
              fontSize:11, fontWeight:600,
              display:"inline-flex", alignItems:"center", gap:5,
            }}
            title="검색 (Ctrl+F) — 노드/엣지/IC/IT 전역 검색"
          >
            🔍 Search
            <span style={{ fontSize:9, opacity:0.7, background:"#0f172a", padding:"1px 4px", borderRadius:3, marginLeft:2 }}>
              Ctrl+F
            </span>
          </button>

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

          {/* v10.3+: Interface 모드 토글 — IT 라인 강조 / 다른 라인 페이드 */}
          <button onClick={()=>setInterfaceMode(v=>!v)}
            style={{
              background: interfaceMode ? "#7c3aed" : "#1e293b",
              color: interfaceMode ? "#fff" : "#c4b5fd",
              border: `1px solid ${interfaceMode ? "#a78bfa" : "#475569"}`,
              borderRadius:5, padding:"3px 12px", cursor:"pointer",
              fontSize:11, fontWeight:700,
              transition: "all 0.15s ease",
            }}
            title="Interface 모드: IT 라인만 선명하게, 다른 라인은 페이드"
          >
            🔗 Interface 모드 {interfaceMode && "ON"}
          </button>

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
              className={interfaceMode ? "mbse-interface-mode" : ""}
              nodes={nodes} edges={edges}
              onNodesChange={onNodesChange} onEdgesChange={onEdgesChange}
              onConnect={onConnect}
              // ── IT 라인 끝점 재연결 (드래그로 다른 IT 포트로 이동) ──
              onReconnect={onReconnectIT}
              onReconnectStart={onReconnectStart}
              onReconnectEnd={onReconnectEnd}
              edgesReconnectable={true}
              reconnectRadius={24}
              // 연결 시 핸들 스냅 반경 확대 — 핸들 근처(36px)에서 자동 스냅
              connectionRadius={36}
              onNodeClick={onNodeClick} onEdgeClick={onEdgeClick} onPaneClick={onPaneClick}
              onNodeDragStart={onNodeDragStart}
              onNodeDrag={onNodeDrag}
              onNodeDragStop={onNodeDragStop}
              nodeTypes={nodeTypes} edgeTypes={edgeTypes}
              connectionMode={ConnectionMode.Loose}
              connectionLineType="straight"
              connectionLineStyle={{ stroke:"#0d9488",strokeWidth:2,strokeDasharray:"4 2" }}
              defaultEdgeOptions={{ type:"pipe" }}
              // ── 성능 최적화: viewport 밖 노드 렌더링 안 함 ──
              // 화면에 안 보이는 노드/엣지는 DOM 에서 제외 → 대규모 모델에서 부드러움
              onlyRenderVisibleElements={true}
              // ── 노드 선택 시 자동 z-index 상승 비활성화 ──
              // Area를 클릭해도 내부 Equipment가 위에 머물러 클릭 가능
              elevateNodesOnSelect={false}
              elevateEdgesOnSelect={false}
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
      {modal&&<ConnModal defaultType={modalDefault} canBeIT={canBeIT} onConfirm={confirmConn} onCancel={()=>{ setModal(false); connRef.current=null; setCanBeIT(false); }}/>}

      {/* v10.3: IT Line Definition + Check List 모달 */}
      {itModalEdgeId && (() => {
        const itEdge = edges.find(e => e.id === itModalEdgeId && e.type === "itEdge");
        if (!itEdge) return null;
        const allITEdges = edges.filter(e => e.type === "itEdge");
        return (
          <ITLineModal
            edge={itEdge}
            nodes={nodes}
            allITEdges={allITEdges}
            onSave={(id, newData) => {
              setEdges(es => es.map(e => e.id === id ? { ...e, data: { ...e.data, ...newData } } : e));
              setItModalEdgeId(null);
            }}
            onClose={() => setItModalEdgeId(null)}
            onNavigate={(newId) => setItModalEdgeId(newId)}
            onDelete={(id) => {
              setEdges(es => es.filter(e => e.id !== id));
              setItModalEdgeId(null);
              setSaveMsg("🗑 IT Line 삭제됨");
              setTimeout(()=>setSaveMsg(""), 2500);
            }}
          />
        );
      })()}

      {/* v10.3+: 특정 Area의 IT 목록 팝업 (배지 클릭 시 표시) */}
      {itListAreaId && (() => {
        const area = nodes.find(n => n.id === itListAreaId);
        if (!area) return null;
        const list = edges
          .filter(e => e.type === "itEdge" && (e.source === itListAreaId || e.target === itListAreaId))
          .map(e => {
            const otherId = e.source === itListAreaId ? e.target : e.source;
            const otherNode = nodes.find(n => n.id === otherId);
            return { edge: e, otherName: otherNode?.data?.label || otherId };
          });
        return (
          <div
            onClick={() => setItListAreaId(null)}
            style={{
              position:"fixed", inset:0, background:"rgba(15,23,42,0.5)",
              zIndex:999, display:"flex", alignItems:"center", justifyContent:"center",
            }}
          >
            <div
              onClick={e => e.stopPropagation()}
              style={{
                width:480, maxHeight:"75vh", background:"#fff", borderRadius:10,
                boxShadow:"0 20px 50px rgba(0,0,0,0.3)", overflow:"hidden",
                display:"flex", flexDirection:"column",
              }}
            >
              <div style={{
                padding:"12px 18px", background:"#7c3aed", color:"#fff",
                display:"flex", alignItems:"center", justifyContent:"space-between",
              }}>
                <div>
                  <div style={{ fontSize:11, opacity:0.85 }}>{area.data?.areaType || "Area"}</div>
                  <div style={{ fontSize:15, fontWeight:800 }}>
                    🔗 {area.data?.label || area.id} — IT Line {list.length}개
                  </div>
                </div>
                <button onClick={() => setItListAreaId(null)} style={{
                  background:"rgba(255,255,255,0.25)", border:"none", color:"#fff",
                  padding:"4px 10px", borderRadius:5, cursor:"pointer", fontSize:11, fontWeight:700,
                }}>✕</button>
              </div>
              <div style={{ flex:1, overflowY:"auto", padding:"6px 0" }}>
                {list.length === 0 ? (
                  <div style={{ padding:30, textAlign:"center", color:"#94a3b8", fontSize:12 }}>
                    이 Area에 연결된 IT Line이 없습니다.
                  </div>
                ) : list.map(({edge, otherName}) => {
                  const cl = edge.data?.checklist || {};
                  const total = Object.keys(cl).length || 5;
                  const done = Object.values(cl).filter(itCatHasContent).length;
                  const pct = total > 0 ? done/total : 0;
                  const statusColor = pct === 1 ? "#16a34a" : pct > 0 ? "#f59e0b" : "#94a3b8";
                  return (
                    <div key={edge.id}
                      onClick={() => { setItListAreaId(null); setItModalEdgeId(edge.id); }}
                      style={{
                        padding:"10px 18px", cursor:"pointer", borderBottom:"1px solid #f1f5f9",
                        display:"flex", alignItems:"center", gap:10,
                        transition: "background 0.1s ease",
                      }}
                      onMouseEnter={e=>e.currentTarget.style.background="#f8fafc"}
                      onMouseLeave={e=>e.currentTarget.style.background=""}
                    >
                      <div style={{
                        background:"#7c3aed", color:"#fff", padding:"3px 9px",
                        borderRadius:10, fontSize:11, fontWeight:800, fontFamily:"monospace",
                        border:`2px solid ${statusColor}`,
                        whiteSpace:"nowrap",
                      }}>
                        {edge.data?.itNo || "IT-?"}
                      </div>
                      <div style={{ flex:1, fontSize:12, color:"#1e293b" }}>
                        <div style={{ fontWeight:700 }}>↔ {otherName}</div>
                        {(() => {
                          // 영향인자: 구버전(문자열)·신버전(객체) 모두 text 추출
                          const texts = (edge.data?.influenceFactors||[])
                            .map(f=>typeof f==="string"?f:(f?.text||""))
                            .filter(t=>t.trim());
                          if (texts.length === 0) return null;
                          const joined = texts.join(" · ");
                          return (
                            <div style={{ fontSize:10.5, color:"#64748b", marginTop:2 }}>
                              영향: {joined.slice(0,60)}{joined.length>60 ? "…":""}
                            </div>
                          );
                        })()}
                      </div>
                      <div style={{
                        fontSize:10, color: statusColor, fontWeight:700,
                        background: statusColor+"15", padding:"2px 8px", borderRadius:4,
                        whiteSpace:"nowrap",
                      }}>
                        {done}/{total}
                      </div>
                    </div>
                  );
                })}
              </div>
              <div style={{
                padding:"8px 18px", borderTop:"1px solid #e2e8f0",
                background:"#f8fafc", fontSize:11, color:"#64748b", textAlign:"center",
              }}>
                항목을 클릭하면 상세 편집 화면이 열립니다
              </div>
            </div>
          </div>
        );
      })()}

      {/* v10.4: 전역 검색 모달 (Ctrl+F) */}
      {showSearch && (
        <SearchModal
          nodes={nodes}
          edges={edges}
          onSelect={handleSearchSelect}
          onClose={()=>setShowSearch(false)}
        />
      )}

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
