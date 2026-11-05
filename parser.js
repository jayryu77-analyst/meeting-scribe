/* 일정 파서: 엑셀 행(2차원 배열) 또는 텍스트를 미팅 목록으로 바꿉니다.
   결과: {title, meetings:[{date,start,end,org,contact,tel,loc,brokers:[]}], warnings:[]} */
(function (root) {
  const pad = n => String(n).padStart(2, "0");
  const PHONE = /(0\d{1,2})[-.\s]?(\d{3,4})[-.\s]?(\d{4,5})/;

  // 엑셀 날짜 일련번호 → {y,m,d,H,M}
  function serial(v) {
    const ms = Math.round((v - 25569) * 86400000);
    const dt = new Date(ms);
    return { y: dt.getUTCFullYear(), m: dt.getUTCMonth() + 1, d: dt.getUTCDate(), H: dt.getUTCHours(), M: dt.getUTCMinutes() };
  }
  function normTime(v) {
    if (v === null || v === undefined || v === "") return null;
    if (typeof v === "number") {
      if (v >= 0 && v < 1) { const t = Math.round(v * 1440); return pad(Math.floor(t / 60) % 24) + ":" + pad(t % 60); }
      return null;
    }
    const s = String(v).trim();
    let m = s.match(/^(오전|오후|AM|PM)?\s*(\d{1,2})\s*[:시.]\s*(\d{2})\s*(분)?\s*(AM|PM)?$/i);
    if (!m) m = s.match(/^(오전|오후|AM|PM)?\s*(\d{1,2})\s*[:시]\s*(\d{2})/i);
    if (!m) return null;
    let h = +m[2]; const mi = +m[3];
    const ap = (m[1] || m[5] || "").toUpperCase();
    if ((ap === "오후" || ap === "PM") && h < 12) h += 12;
    if (h > 23 || mi > 59) return null;
    return pad(h) + ":" + pad(mi);
  }
  function normDate(v, year) {
    if (v === null || v === undefined || v === "") return null;
    if (typeof v === "number") {
      if (v > 30000 && v < 80000) { const p = serial(v); return `${p.y}-${pad(p.m)}-${pad(p.d)}`; }
      return null;
    }
    const s = String(v).trim();
    let m = s.match(/(\d{4}|\d{2})\s*[.\-\/년]\s*(\d{1,2})\s*[.\-\/월]\s*(\d{1,2})/);
    if (m) { let y = +m[1]; if (y < 100) y += 2000; return valid(y, +m[2], +m[3]); }
    m = s.match(/(\d{1,2})\s*월\s*(\d{1,2})\s*일/) || s.match(/^\D{0,6}(\d{1,2})\s*\/\s*(\d{1,2})(?!\d)/);
    if (m) return valid(year, +m[1], +m[2]);
    return null;
  }
  function valid(y, mo, d) { return mo >= 1 && mo <= 12 && d >= 1 && d <= 31 ? `${y}-${pad(mo)}-${pad(d)}` : null; }

  function colType(h) {
    h = String(h || "").trim().toLowerCase();
    if (!h) return null;
    if (/소요|duration|분량/.test(h)) return "dur";
    if (/브로커|broker|sales|세일즈|영업/.test(h)) return "broker";
    if (/^end|종료|끝/.test(h)) return "end";
    if (/^start|시작|^시간$|^time$|^일시$/.test(h)) return "start";
    if (/tel|phone|전화|연락처|휴대|mobile/.test(h)) return "tel";
    if (/location|장소|위치|주소|place|venue|address/.test(h)) return "loc";
    if (/contact|담당자|참석자|성함|이름|name|attendee/.test(h)) return "contact";
    if (/기관|방문|회사|운용|investor|institution|client|company|fund/.test(h)) return "org";
    if (/date|날짜|일자/.test(h)) return "date";
    return null;
  }

  function parseRows(rows, opts = {}) {
    const year = opts.year || new Date().getFullYear();
    const out = []; const warnings = [];
    let title = ""; let curDate = null; let map = null;
    for (const raw of rows) {
      const row = (raw || []).map(c => (c === null || c === undefined) ? "" : c);
      const cells = row.map(c => typeof c === "string" ? c.trim() : c);
      const filled = cells.filter(c => c !== "");
      if (!filled.length) continue;
      const joined = filled.map(String).join(" ");

      if (!title) { const t = filled.find(c => typeof c === "string" && /NDR|IR|로드쇼|roadshow|미팅|일정/i.test(c) && !normTime(c)); if (t) title = String(t).replace(/\s+/g, " ").trim(); }

      // 헤더 행
      const types = cells.map(colType);
      if (types.filter(Boolean).length >= 2 && types.includes("start") || types.filter(t => t && t !== "dur").length >= 3) {
        if (!cells.some(c => normTime(c))) { map = {}; types.forEach((t, i) => { if (t && map[t] === undefined) map[t] = i; }); continue; }
      }

      // 날짜만 있는 행 (예: 26.11.05(목))
      const timeCells = cells.map(normTime);
      const hasTime = timeCells.some(Boolean);
      if (!hasTime) {
        const d = filled.map(c => normDate(c, year)).find(Boolean);
        if (d && filled.length <= 3) { curDate = d; continue; }
        if (d && /일시|date|날짜/i.test(joined)) { curDate = d; continue; }
        continue;
      }
      if (/점심|lunch|오찬|이동|break/i.test(joined) && filled.filter(c => typeof c === "string" && !normTime(c)).length <= 2 && !PHONE.test(joined)) continue;

      let m = { date: null, start: null, end: null, org: "", contact: "", tel: "", loc: "", brokers: [] };
      if (map && map.start !== undefined && normTime(cells[map.start])) {
        const g = k => map[k] !== undefined ? cells[map[k]] : "";
        m.start = normTime(g("start")); m.end = normTime(g("end"));
        m.org = String(g("org") || "").trim(); m.contact = String(g("contact") || "").trim();
        m.tel = String(g("tel") || "").trim(); m.loc = String(g("loc") || "").trim();
        m.brokers = splitBrokers(g("broker"));
        if (map.date !== undefined) m.date = normDate(g("date"), year);
      } else {
        // 헤더 없이: 시간 2개 → 시작/종료, 전화번호, 나머지 글자 순서대로 기관/담당자/장소/브로커
        const idx = timeCells.map((t, i) => t ? i : -1).filter(i => i >= 0);
        m.start = timeCells[idx[0]]; m.end = idx[1] !== undefined ? timeCells[idx[1]] : null;
        const rest = cells.filter((c, i) => !timeCells[i] && c !== "" && typeof c !== "number");
        const texts = [];
        for (const c of rest) {
          const s = String(c);
          if (!m.tel && PHONE.test(s) && s.replace(/[\d\-.\s()+]/g, "").length < 3) { m.tel = s; continue; }
          const d = normDate(s, year); if (d && s.length < 14) { m.date = d; continue; }
          texts.push(s);
        }
        [m.org = "", m.contact = "", m.loc = ""] = texts;
        m.brokers = splitBrokers(texts[3] || "");
        if (texts.length === 1) { const sp = texts[0].split(/\s+/); m.org = sp[0] || ""; m.contact = sp[1] || ""; m.loc = sp.slice(2).join(" "); }
      }
      // 소요시간 칸이 종료로 잡힌 경우 보정 (예: 0:50)
      if (m.end && m.start && m.end < m.start && /^00:/.test(m.end)) {
        const [h, mi] = m.start.split(":").map(Number); const add = +m.end.slice(3);
        const t = h * 60 + mi + add; m.end = pad(Math.floor(t / 60)) + ":" + pad(t % 60);
      }
      m.date = m.date || curDate;
      if (!m.date) { warnings.push(`날짜를 찾지 못한 행: ${joined.slice(0, 40)}`); continue; }
      if (!m.end) { const [h, mi] = m.start.split(":").map(Number); const t = h * 60 + mi + 50; m.end = pad(Math.floor(t / 60)) + ":" + pad(t % 60); }
      if (!m.org) m.org = "(기관명 없음)";
      out.push(m);
    }
    out.sort((a, b) => (a.date + a.start).localeCompare(b.date + b.start));
    return { title, meetings: out, warnings };
  }
  function splitBrokers(v) { return String(v || "").split(/[\/,·&]|\s{2,}/).map(s => s.trim()).filter(Boolean); }

  function parseText(text, opts = {}) {
    const rows = String(text || "").replace(/\r/g, "").split("\n").map(line => {
      if (line.includes("\t")) return line.split("\t");
      if (/,/.test(line) && line.split(",").length >= 3) return line.split(",");
      if (/\s{2,}/.test(line.trim())) return line.trim().split(/\s{2,}/);
      // "09:10~10:00 베어링자산 김재민 매니저 ..." 형태
      const m = line.trim().match(/^(\d{1,2}:\d{2})\s*[-~–]\s*(\d{1,2}:\d{2})\s+(.+)$/);
      if (m) {
        let rest = m[3]; let tel = "";
        const p = rest.match(PHONE); if (p) { tel = p[0]; rest = rest.replace(p[0], " "); }
        const toks = rest.trim().split(/\s+/);
        let contact = toks[1] || "", locStart = 2;
        if (toks[2] && /^(매니저|상무|전무|이사|부장|차장|과장|대리|팀장|책임|수석|본부장|실장|대표|선임|사원|위원|센터장|CIO|PM)$/.test(toks[2])) { contact += " " + toks[2]; locStart = 3; }
        return [m[1], m[2], toks[0] || "", contact, tel, toks.slice(locStart).join(" ")];
      }
      return [line];
    });
    return parseRows(rows, opts);
  }
  const api = { parseRows, parseText, normTime, normDate };
  if (typeof module !== "undefined") module.exports = api; else root.SchedParser = api;
})(this);
