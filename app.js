/* NDR 미팅노트 — 일정 업로드 · 받아쓰기 · 메모 · AI 정리 (모든 데이터는 이 기기 브라우저에 저장) */
"use strict";
const $ = id => document.getElementById(id);
const pad = n => String(n).padStart(2, "0");
const esc = s => String(s ?? "").replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
const fmt = ms => { const s = Math.max(0, Math.floor(ms / 1000)); return pad(Math.floor(s / 60)) + ":" + pad(s % 60); };
const kst = (d, t) => new Date(`${d}T${t}:00+09:00`);
const todayStr = () => new Date().toLocaleDateString("sv-SE", { timeZone: "Asia/Seoul" });
const hhmm = () => new Date().toLocaleTimeString("ko-KR", { hour: "2-digit", minute: "2-digit", hour12: false, timeZone: "Asia/Seoul" });
const dow = d => "일월화수목금토"[new Date(d + "T12:00:00+09:00").getDay()];
const md = d => `${+d.slice(5, 7)}/${+d.slice(8)}`;

/* ---------- 저장 ---------- */
let photoFile = null;
const DB_KEY = "ndrnote-v1", SET_KEY = "ndrnote-settings";
function loadDB() { try { const v = JSON.parse(localStorage.getItem(DB_KEY) || "null"); if (v && v.schedules) return v; } catch (e) { } return { schedules: [], notes: {}, sums: {}, cur: null }; }
let S = loadDB();
let saveT = null;
function saveNow() { try { localStorage.setItem(DB_KEY, JSON.stringify(S)); } catch (e) { msg("저장 공간이 부족합니다. 백업 후 오래된 일정을 지워주세요."); } }
function saveSoon() { clearTimeout(saveT); saveT = setTimeout(saveNow, 700); }
window.addEventListener("pagehide", saveNow);
document.addEventListener("visibilitychange", () => { if (document.visibilityState === "hidden") saveNow(); });
function settings() { try { return JSON.parse(localStorage.getItem(SET_KEY) || "{}"); } catch (e) { return {}; } }

const cur = () => S.schedules.find(s => s.id === S.cur) || null;
const nkey = (sid, mid) => sid + "|" + mid;
function note(sid, mid) { const k = nkey(sid, mid); return S.notes[k] || (S.notes[k] = { tr: "", note: "", at: 0 }); }

function msg(t, ok) { const b = $("msg"); b.hidden = !t; b.textContent = t || ""; b.className = "warnbox" + (ok ? " ok" : ""); if (t) { clearTimeout(msg.t); msg.t = setTimeout(() => b.hidden = true, 6000); } }

/* ---------- 화면 전환 ---------- */
function show(v) {
  if (v !== "sched") $("inbox").hidden = true;
  $("importView").hidden = v !== "import"; $("schedView").hidden = v !== "sched"; $("setView").hidden = v !== "set";
  $("cancelImport").hidden = !S.schedules.length;
  fillSel();
}
function fillSel() {
  const sel = $("schedSel");
  sel.hidden = !S.schedules.length;
  sel.innerHTML = S.schedules.map(s => `<option value="${s.id}" ${s.id === S.cur ? "selected" : ""}>${esc(s.title)}</option>`).join("");
}
$("schedSel").onchange = e => { S.cur = e.target.value; saveNow(); ui.day = null; ui.broker = "전체"; renderSched(); };
$("newBtn").onclick = () => { resetImport(); show("import"); };
$("cancelImport").onclick = () => { show("sched"); renderSched(); };
$("setBtn").onclick = () => { const st = settings(); $("keyIn").value = st.key || ""; $("modelIn").value = st.model || ""; $("sumUrlIn").value = sumUrl(); $("recAppIn").value = st.recApp || "browser"; show("set"); };
$("closeSet").onclick = () => S.schedules.length ? (show("sched"), renderSched()) : show("import");
$("saveSet").onclick = () => {
  try { localStorage.setItem(SET_KEY, JSON.stringify({ key: $("keyIn").value.trim(), model: $("modelIn").value.trim(), sumUrl: $("sumUrlIn").value.trim(), recApp: $("recAppIn").value })); } catch (e) { }
  msg("설정을 저장했습니다.", true); $("closeSet").onclick();
};

/* ---------- 일정 가져오기 ---------- */
let parsed = null;
function resetImport() { parsed = null; $("preview").hidden = true; $("pasteIn").value = ""; $("fileIn").value = ""; photoFile = null; $("photoBox").hidden = true; }
function loadXLSX() {
  if (window.XLSX) return Promise.resolve();
  return new Promise((ok, no) => { const s = document.createElement("script"); s.src = "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"; s.onload = ok; s.onerror = () => no(new Error("xlsx")); document.head.appendChild(s); });
}
$("fileIn").onchange = async e => {
  const f = e.target.files[0]; if (!f) return;
  const base = f.name.replace(/\.[^.]+$/, "");
  try {
    if (/\.(xlsx|xls)$/i.test(f.name)) {
      await loadXLSX();
      const wb = XLSX.read(await f.arrayBuffer(), { type: "array" });
      let rows = [];
      wb.SheetNames.forEach(n => { rows = rows.concat(XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: "" })); rows.push([]); });
      preview(SchedParser.parseRows(rows, { year: new Date().getFullYear() }), base);
    } else {
      preview(SchedParser.parseText(await f.text(), { year: new Date().getFullYear() }), base);
    }
  } catch (err) { msg(err.message === "xlsx" ? "엑셀 읽기 도구를 불러오지 못했습니다. 인터넷 연결을 확인하거나, 엑셀 표를 복사해 붙여넣어 주세요." : "파일을 읽지 못했습니다. 엑셀 표를 복사해 붙여넣어 보세요."); }
};
$("parseBtn").onclick = () => {
  const t = $("pasteIn").value.replace(/```[a-z]*\n?/gi, "").trim();
  if (!t) { msg("파일을 올리거나 일정을 붙여넣어 주세요."); return; }
  preview(SchedParser.parseText(t, { year: new Date().getFullYear() }), "");
};
function preview(r, fallbackTitle) {
  parsed = r;
  if (!r.meetings.length) { $("preview").hidden = true; msg("미팅을 찾지 못했습니다. 시간(09:10)과 날짜(26.11.05 또는 11/5)가 들어 있는지 확인해 주세요."); return; }
  $("preview").hidden = false;
  $("titleIn").value = r.title || fallbackTitle || "새 NDR 일정";
  let last = "", html = "";
  r.meetings.forEach(m => {
    if (m.date !== last) { html += `<div class="pd">${md(m.date)}(${dow(m.date)})</div>`; last = m.date; }
    html += `<div class="pm"><span>${m.start}~${m.end}</span><span><b>${esc(m.org)}</b> ${esc(m.contact)} ${m.loc ? "· " + esc(m.loc) : ""} ${m.brokers.length ? "· " + esc(m.brokers.join("/")) : ""}</span></div>`;
  });
  $("prevList").innerHTML = html;
  const days = new Set(r.meetings.map(m => m.date)).size;
  $("prevWarn").textContent = `${days}일 · 미팅 ${r.meetings.length}건을 찾았습니다.` + (r.warnings.length ? ` 건너뛴 줄 ${r.warnings.length}개: ${r.warnings.slice(0, 2).join(" / ")}` : "");
  $("preview").scrollIntoView({ behavior: "smooth", block: "start" });
}
$("toTsv").onclick = () => {
  if (!parsed) return;
  $("pasteIn").value = "TITLE\t" + $("titleIn").value + "\n" + parsed.meetings.map(m => [m.date, m.start, m.end, m.org, m.contact, m.tel, m.loc, m.brokers.join("/")].join("\t")).join("\n");
  $("pasteIn").focus(); $("pasteIn").scrollIntoView({ behavior: "smooth", block: "center" });
  msg("한 줄에 한 미팅입니다(날짜·시작·종료·기관·담당자·전화·장소·브로커, 탭 구분). 틀린 글자를 고친 뒤 '읽어오기'를 누르세요.", true);
};
$("createBtn").onclick = () => {
  if (!parsed) return;
  const id = "s" + Date.now().toString(36);
  const meetings = parsed.meetings.map((m, i) => ({ id: "m" + i.toString(36) + Math.random().toString(36).slice(2, 5), ...m }));
  S.schedules.unshift({ id, title: $("titleIn").value.trim() || "새 NDR 일정", meetings, createdAt: Date.now() });
  S.cur = id; saveNow(); resetImport();
  ui.day = null; ui.broker = "전체"; ui.open = new Set();
  show("sched"); renderSched(); msg("일정을 만들었습니다.", true);
};


/* ---------- 사진으로 일정 만들기 ---------- */
const PHOTO_PROMPT = () => `이 사진은 기관투자자 미팅(NDR) 일정표입니다. 아래 형식의 탭 구분 텍스트로만 답하세요. 설명이나 코드블록 없이 결과만 쓰세요.
첫 줄: TITLE<탭>일정 이름
다음 줄부터 미팅 한 줄씩: 날짜(YYYY-MM-DD)<탭>시작(HH:MM)<탭>종료(HH:MM)<탭>방문기관<탭>담당자(이름 직함)<탭>전화<탭>장소<탭>담당브로커(여러 명이면 /로 구분)
점심 등 미팅이 아닌 줄은 빼세요. 연도가 없으면 ${new Date().getFullYear()}년으로 쓰세요. 비어 있는 칸은 비워 두고 탭은 유지하세요.`;
function loadScript(src) { return new Promise((ok, no) => { const s = document.createElement("script"); s.src = src; s.onload = ok; s.onerror = () => no(new Error("script")); document.head.appendChild(s); }); }
async function imgCanvas(file, maxSide, forOcr) {
  const bmp = await createImageBitmap(file);
  const k = Math.min(forOcr ? 2 : 1, maxSide / Math.max(bmp.width, bmp.height));
  const c = document.createElement("canvas"); c.width = Math.round(bmp.width * k); c.height = Math.round(bmp.height * k);
  const g = c.getContext("2d"); g.drawImage(bmp, 0, 0, c.width, c.height);
  if (forOcr) { // 흑백 + 어두운 배경이면 반전 + 대비 높이기
    const im = g.getImageData(0, 0, c.width, c.height), d = im.data; let sum = 0;
    for (let i = 0; i < d.length; i += 4) { const y = 0.299 * d[i] + 0.587 * d[i + 1] + 0.114 * d[i + 2]; d[i] = d[i + 1] = d[i + 2] = y; sum += y; }
    const inv = sum / (d.length / 4) < 115;
    for (let i = 0; i < d.length; i += 4) { let y = d[i]; if (inv) y = 255 - y; y = y < 150 ? Math.max(0, y - 40) : 255; d[i] = d[i + 1] = d[i + 2] = y; }
    g.putImageData(im, 0, 0);
  }
  return c;
}
async function readPhoto() {
  if (!photoFile) return;
  const st = $("photoSt");
  if (settings().key) {
    st.textContent = "AI가 사진을 읽는 중… (10~30초)";
    try {
      const c = await imgCanvas(photoFile, 2000, false);
      const b64 = c.toDataURL("image/jpeg", 0.9).split(",")[1];
      const text = await askAI([{ type: "image", source: { type: "base64", media_type: "image/jpeg", data: b64 } }, { type: "text", text: PHOTO_PROMPT() }]);
      $("pasteIn").value = text; st.textContent = "AI가 읽었습니다. 미리보기를 확인하세요."; $("parseBtn").click(); return;
    } catch (e) { st.textContent = "AI로 읽지 못해 이 기기에서 읽습니다… (" + e.message + ")"; }
  }
  try {
    st.textContent = "글자 인식 도구를 불러오는 중… (처음 한 번은 수십 MB를 받습니다. 와이파이 권장)";
    if (!window.Tesseract) await loadScript("https://cdn.jsdelivr.net/npm/tesseract.js@5.1.1/dist/tesseract.min.js");
    const c = await imgCanvas(photoFile, 2600, true);
    const w = await Tesseract.createWorker("kor+eng", 1, { logger: m => { if (m.status === "recognizing text") st.textContent = `사진 속 글자 읽는 중… ${Math.round(m.progress * 100)}%`; else if (/load|initializ/.test(m.status)) st.textContent = "글자 인식 준비 중…"; } });
    await w.setParameters({ tessedit_pageseg_mode: "6", preserve_interword_spaces: "1" });
    const r = await w.recognize(c); await w.terminate();
    const text = r.data.text.replace(/[|｜]/g, " ").replace(/[ \t]{2,}/g, " ");
    $("pasteIn").value = text;
    st.textContent = "이 기기에서 읽었습니다. 한글 이름은 틀릴 수 있으니 미리보기를 확인하세요. 더 정확하게 하려면 'Claude 앱에서 읽기'를 쓰세요.";
    $("parseBtn").click();
  } catch (e) { st.textContent = "사진을 읽지 못했습니다. 'Claude 앱에서 읽기'를 써보세요."; }
}
$("photoIn").onchange = e => {
  photoFile = e.target.files[0]; e.target.value = ""; if (!photoFile) return;
  $("photoBox").hidden = false; $("photoPrev").src = URL.createObjectURL(photoFile); $("preview").hidden = true;
  readPhoto();
};
$("photoAgain").onclick = readPhoto;
$("photoClaude").onclick = async e => {
  if (!photoFile) return; const btn = e.currentTarget, st = $("photoSt");
  const mobile = /Android|iPhone|iPad/i.test(navigator.userAgent);
  if (mobile && navigator.canShare && navigator.canShare({ files: [photoFile] })) {
    try { await navigator.share({ files: [photoFile], text: PHOTO_PROMPT(), title: "일정표 읽기" }); st.textContent = "Claude 앱을 골라 보내세요. Claude가 준 결과를 아래 칸에 붙여넣고 '읽어오기'를 누르세요."; return; }
    catch (err) { if (err && err.name === "AbortError") return; }
  }
  await copy(PHOTO_PROMPT(), btn, "Claude 앱에서 읽기 (더 정확)");
  st.textContent = "요청문을 복사했습니다. Claude 앱에서 이 사진을 첨부하고 요청문을 붙여넣은 뒤, 결과를 아래 칸에 붙여넣으세요.";
};

/* ---------- 일정 화면 ---------- */
const ui = { day: null, broker: "전체", open: new Set(), first: true };
function dates(s) { return [...new Set(s.meetings.map(m => m.date))].sort(); }
function pickDay(s) { const ds = dates(s), t = todayStr(); return ds.includes(t) ? t : (ds.find(d => d > t) || ds[ds.length - 1]); }
function renderSched() {
  const s = cur(); if (!s) { show("import"); return; }
  show("sched");
  const ds = dates(s);
  if (!ui.day || !ds.includes(ui.day)) ui.day = pickDay(s);
  $("sTitle").innerHTML = esc(s.title).replace(/^(.+?)(\s*NDR)/i, "<span>$1</span>$2");
  $("sRange").textContent = `${ds[0].replace(/-/g, ".")} – ${md(ds[ds.length - 1])} · ${ds.length}일`;
  // 다음 미팅
  const now = new Date(); const nx = s.meetings.find(m => kst(m.date, m.end) > now);
  $("stats").innerHTML = `<div class="stat"><b>${ds.length}</b><small>일정 (일)</small></div><div class="stat"><b>${s.meetings.length}</b><small>미팅</small></div>
    <div class="stat"><b style="font-size:15px">${nx ? md(nx.date) + " " + nx.start : "완료"}</b><small>${nx ? esc(nx.org) : "전 일정 종료"}</small></div>`;
  $("tabs").innerHTML = ds.map(d => { const n = s.meetings.filter(m => m.date === d).length; return `<button class="tab" type="button" role="tab" data-d="${d}" aria-selected="${d === ui.day}"><span class="d">${md(d)}(${dow(d)})</span><span class="n">${n}개 미팅</span></button>`; }).join("");
  $("tabs").querySelectorAll(".tab").forEach(b => b.onclick = () => { ui.day = b.dataset.d; renderSched(); });
  const brokers = [...new Set(s.meetings.flatMap(m => m.brokers))];
  $("filters").innerHTML = brokers.length ? ["전체", ...brokers].map(b => { const c = b === "전체" ? s.meetings.length : s.meetings.filter(m => m.brokers.includes(b)).length; return `<button class="chip" type="button" data-b="${esc(b)}" aria-pressed="${b === ui.broker}">${esc(b)} ${c}</button>`; }).join("") : "";
  $("filters").querySelectorAll(".chip").forEach(c => c.onclick = () => { ui.broker = c.dataset.b; renderSched(); });
  renderTimeline(s); paintSum();
}
function renderTimeline(s) {
  const tl = $("timeline"); const now = new Date();
  const ms = s.meetings.filter(m => m.date === ui.day && (ui.broker === "전체" || m.brokers.includes(ui.broker)));
  if (!ms.length) { tl.innerHTML = `<div class="empty">${md(ui.day)}(${dow(ui.day)})에는 ${esc(ui.broker)} 담당 미팅이 없습니다.</div>`; return; }
  let html = ""; let prev = null;
  for (const m of ms) {
    if (prev && ui.broker === "전체" && prev.end <= "12:30" && m.start >= "12:30" && m.start > prev.end) html += `<div class="lunch">점심</div>`;
    if (ui.first && now >= kst(m.date, m.start) && now < kst(m.date, m.end)) ui.open.add(m.id);
    html += card(s, m); prev = m;
  }
  ui.first = false;
  tl.innerHTML = html; wire(s); refreshState(); syncBadges();
}
function card(s, m) {
  const n = note(s.id, m.id), open = ui.open.has(m.id);
  const loc = m.loc ? `<span>${esc(m.loc)} <a class="b sm" href="https://map.naver.com/p/search/${encodeURIComponent(m.loc.replace(/\s?\d+층$/, ""))}" target="_blank" rel="noopener">지도</a></span>` : "";
  const tel = m.tel ? `<span><span class="num">${esc(m.tel)}</span> <button class="b sm" type="button" data-copy="${esc(m.tel)}">복사</button></span>` : "";
  return `<article class="mtg ${open ? "open" : ""}" data-id="${m.id}" data-d="${m.date}" data-s="${m.start}" data-e="${m.end}">
    <div class="time"><span class="s">${m.start}</span><span class="e">~ ${m.end}</span></div>
    <div class="body">
      <div><span class="org">${esc(m.org)}</span><span class="badge" hidden>진행 중</span></div>
      <div class="info">${m.contact ? `<strong>${esc(m.contact)}</strong>` : ""}${tel}${loc}</div>
      ${m.brokers.length ? `<div class="bks">${m.brokers.map(b => `<span class="bk">${esc(b)}</span>`).join("")}</div>` : ""}
      <div class="notebar"><button class="b" type="button" data-toggle="${m.id}" aria-expanded="${open}">기록</button><span class="nb" data-badge="${m.id}"></span></div>
      <div class="panel" data-panel="${m.id}" ${open ? "" : "hidden"}>
        <div class="row"><button class="b rec" type="button" data-rec="${m.id}">받아쓰기 시작</button><button class="b sm" type="button" data-alt="${m.id}"></button></div>
        <div class="st" data-rst="${m.id}">${extStatus(n)}</div>
        <h4>녹취 (자동 받아쓰기 · 다른 녹음 앱 결과를 붙여넣어도 됩니다)</h4>
        <textarea class="tr" id="tr-${m.id}" data-tr="${m.id}" rows="5" placeholder="받아쓰기를 시작하면 여기에 [분:초]와 함께 적힙니다.">${esc(n.tr)}</textarea>
        <div class="st" id="int-${m.id}"></div>
        <div class="row"><button class="b pri" type="button" data-ai="${m.id}">AI로 노트 정리</button><span class="st" data-ast="${m.id}"></span></div>
        <h4>메모</h4>
        <textarea class="note" id="note-${m.id}" data-note="${m.id}" rows="6" placeholder="질문, 회사 답변, 투자자 반응, 후속 확인할 것">${esc(n.note)}</textarea>
      </div>
    </div></article>`;
}
function wire(s) {
  const tl = $("timeline");
  tl.querySelectorAll("[data-toggle]").forEach(b => b.onclick = () => {
    const id = b.dataset.toggle, p = tl.querySelector(`[data-panel="${id}"]`), open = p.hidden;
    p.hidden = !open; b.setAttribute("aria-expanded", open); b.closest(".mtg").classList.toggle("open", open);
    open ? ui.open.add(id) : ui.open.delete(id);
  });
  tl.querySelectorAll("[data-tr]").forEach(t => t.oninput = () => { const n = note(s.id, t.dataset.tr); n.tr = t.value; n.at = Date.now(); saveSoon(); syncBadges(); });
  tl.querySelectorAll("[data-note]").forEach(t => t.oninput = () => { const n = note(s.id, t.dataset.note); n.note = t.value; n.at = Date.now(); saveSoon(); syncBadges(); });
  tl.querySelectorAll("[data-rec]").forEach(b => b.onclick = () => (R.on && R.mid === b.dataset.rec) ? stopRec() : (recApp() === "browser" ? startRec(s.id, b.dataset.rec) : startExternal(s.id, b.dataset.rec, recApp())));
  tl.querySelectorAll("[data-alt]").forEach(b => b.onclick = () => recApp() === "browser" ? startExternal(s.id, b.dataset.alt, "clova") : ((R.on && R.mid === b.dataset.alt) ? stopRec() : startRec(s.id, b.dataset.alt)));
  tl.querySelectorAll("[data-ai]").forEach(b => b.onclick = () => noteAI(s.id, b.dataset.ai, b));
  tl.querySelectorAll("[data-copy]").forEach(b => b.onclick = () => copy(b.dataset.copy, b));
  paintRecButtons();
}
function refreshState() {
  const now = new Date();
  document.querySelectorAll(".mtg").forEach(el => {
    const a = kst(el.dataset.d, el.dataset.s), z = kst(el.dataset.d, el.dataset.e);
    const isNow = now >= a && now < z; el.classList.toggle("now", isNow); el.classList.toggle("done", now >= z);
    el.querySelector(".badge").hidden = !isNow;
  });
}
function syncBadges() {
  const s = cur(); if (!s) return;
  document.querySelectorAll("[data-badge]").forEach(b => {
    const n = note(s.id, b.dataset.badge), p = [];
    if (n.tr.trim()) p.push("녹취 있음"); if (n.note.trim()) p.push("메모 있음");
    b.textContent = p.join(" · ");
  });
}
async function copy(text, btn, label) {
  const l = label || btn.textContent;
  try { await navigator.clipboard.writeText(text); btn.textContent = "복사됨"; }
  catch (e) { btn.textContent = "복사 실패"; }
  setTimeout(() => btn.textContent = l, 1500);
}

/* ---------- 받아쓰기 ---------- */
const SR = window.SpeechRecognition || window.webkitSpeechRecognition;
const R = { on: false, sid: null, mid: null, rec: null, startAt: 0, interim: "", lastHeard: 0, restarts: 0, wake: null, tick: null };
function paintRecButtons() {
  document.querySelectorAll("[data-rec]").forEach(b => {
    const me = R.on && R.mid === b.dataset.rec;
    const app = recApp();
    b.textContent = me ? "받아쓰기 정지" : app === "browser" ? "받아쓰기 시작" : REC_APPS[app].name + "로 녹음";
    b.disabled = app === "browser" && !SR;
    const alt = document.querySelector(`[data-alt="${b.dataset.rec}"]`);
    if (alt) { alt.textContent = app === "browser" ? "클로바노트로 녹음" : (me ? "받아쓰기 정지" : "앱 안에서 받아쓰기"); alt.hidden = app !== "browser" && !SR; }
    const t = $("tr-" + b.dataset.rec); if (t) t.readOnly = me;
  });
  if (!SR && recApp() === "browser") document.querySelectorAll("[data-rst]").forEach(e => { if (!e.textContent) e.textContent = "이 브라우저는 받아쓰기를 지원하지 않습니다. 설정에서 녹음 방식을 클로바노트로 바꾸거나, 갤럭시는 크롬·아이폰은 사파리로 열어주세요."; });
}
function rstatus(t) { const e = document.querySelector(`[data-rst="${R.mid}"]`); if (e) e.textContent = t; }
async function startRec(sid, mid) {
  if (!SR) return;
  if (R.on) stopRec();
  const s = S.schedules.find(x => x.id === sid), m = s.meetings.find(x => x.id === mid);
  const n = note(sid, mid);
  n.tr = (n.tr.trim() ? n.tr.replace(/\s+$/, "") + "\n\n" : "") + `── 받아쓰기 ${hhmm()} ──\n`;
  Object.assign(R, { on: true, sid, mid, startAt: Date.now(), interim: "", lastHeard: Date.now(), restarts: 0 });
  const rec = new SR(); rec.lang = "ko-KR"; rec.continuous = true; rec.interimResults = true; R.rec = rec;
  rec.onresult = e => {
    let interim = "";
    for (let i = e.resultIndex; i < e.results.length; i++) {
      const t = e.results[i][0].transcript.trim();
      if (e.results[i].isFinal) { if (t) { n.tr += `[${fmt(Date.now() - R.startAt)}] ${t}\n`; n.at = Date.now(); saveSoon(); } }
      else interim += t + " ";
    }
    R.interim = interim; R.lastHeard = Date.now(); R.restarts = 0;
    const ta = $("tr-" + mid); if (ta) { ta.value = n.tr; ta.scrollTop = ta.scrollHeight; }
    const ie = $("int-" + mid); if (ie) ie.textContent = interim ? "듣는 중: " + interim : "";
    rstatus("받아쓰는 중"); syncBadges();
  };
  rec.onerror = e => {
    if (e.error === "not-allowed" || e.error === "service-not-allowed") { stopRec(); msg("마이크 권한이 거부되었습니다. 브라우저 설정에서 이 사이트의 마이크를 허용해 주세요."); }
    else if (e.error === "network") rstatus("네트워크가 불안정합니다. 다시 연결 중…");
  };
  rec.onend = () => {
    if (!R.on || R.rec !== rec) return;
    if (++R.restarts > 8) { rstatus("받아쓰기가 멈췄습니다. 정지 후 다시 시작해 주세요."); return; }
    setTimeout(() => { if (R.on && R.rec === rec) { try { rec.start(); } catch (err) { rstatus("다시 시작하지 못했습니다. 정지 후 다시 눌러주세요."); } } }, 80);
  };
  try { rec.start(); } catch (e) { R.on = false; msg("받아쓰기를 시작하지 못했습니다. 페이지를 새로고침한 뒤 다시 시도해 주세요."); return; }
  try { R.wake = await navigator.wakeLock?.request("screen"); } catch (e) { }
  $("recBar").hidden = false; $("recLabel").textContent = `${m.org} 받아쓰는 중`;
  clearInterval(R.tick);
  R.tick = setInterval(() => { $("recClock").textContent = fmt(Date.now() - R.startAt); if (Date.now() - R.lastHeard > 45000) rstatus("45초 넘게 인식된 말이 없습니다. 휴대폰을 화자 가까이 두세요."); }, 1000);
  const ta = $("tr-" + mid); if (ta) ta.value = n.tr;
  rstatus("받아쓰는 중 · 화면을 켜 두세요"); paintRecButtons();
}
function stopRec() {
  if (!R.on) return;
  const n = note(R.sid, R.mid);
  if (R.interim.trim()) n.tr += `[${fmt(Date.now() - R.startAt)}] ${R.interim.trim()}\n`;
  R.on = false; try { R.rec.stop(); } catch (e) { }
  clearInterval(R.tick); try { R.wake && R.wake.release(); } catch (e) { } R.wake = null;
  n.at = Date.now(); saveNow();
  const ta = $("tr-" + R.mid); if (ta) ta.value = n.tr;
  const ie = $("int-" + R.mid); if (ie) ie.textContent = "";
  rstatus("저장했습니다. 'AI로 노트 정리'로 주요 내용을 뽑을 수 있습니다.");
  $("recBar").hidden = true; paintRecButtons(); syncBadges();
}
$("recStop").onclick = stopRec;
document.addEventListener("visibilitychange", async () => {
  if (document.visibilityState === "visible" && R.on) { try { R.wake = await navigator.wakeLock?.request("screen"); } catch (e) { } try { R.rec.start(); } catch (e) { } }
});

/* ---------- AI 정리 ---------- */
async function askAI(prompt) {
  const st = settings();
  if (!st.key) { const e = new Error("nokey"); e.code = "nokey"; throw e; }
  const r = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: { "content-type": "application/json", "x-api-key": st.key, "anthropic-version": "2023-06-01", "anthropic-dangerous-direct-browser-access": "true" },
    body: JSON.stringify({ model: st.model || "claude-sonnet-5-5", max_tokens: 2500, messages: [{ role: "user", content: prompt }] })
  });
  const j = await r.json().catch(() => ({}));
  if (!r.ok) { const e = new Error(j.error?.message || ("HTTP " + r.status)); e.code = "api"; throw e; }
  return (j.content || []).map(c => c.text || "").join("").trim();
}
function meetingPrompt(s, m, n) {
  return `셀사이드 애널리스트가 ${s.title}에서 기관투자자 ${m.org}${m.contact ? "(" + m.contact + ")" : ""}와 한 미팅 기록입니다.
녹취는 음성 인식 결과라 오류가 섞여 있을 수 있습니다. 주요 내용 위주로 미팅 노트를 만드세요. 기록에 없는 내용은 지어내지 마세요.
한국어 일반 텍스트, 마크다운 기호(#, **) 없이 아래 형식으로 짧게:
[투자자 질문]
- …
[회사 답변 핵심(숫자·가이던스 포함)]
- …
[투자자 반응·톤]
- …
[후속 확인]
- …

[녹취]
${(n.tr || "").slice(0, 60000) || "(없음)"}

[애널리스트 메모]
${n.note || "(없음)"}`;
}
async function noteAI(sid, mid, btn) {
  const s = S.schedules.find(x => x.id === sid), m = s.meetings.find(x => x.id === mid), n = note(sid, mid);
  const st = document.querySelector(`[data-ast="${mid}"]`) || {};
  if (!n.tr.trim() && !n.note.trim()) { st.textContent = "정리할 녹취나 메모가 없습니다."; return false; }
  const prompt = meetingPrompt(s, m, n);
  btn.disabled = true; st.textContent = "정리하는 중…";
  try {
    const text = await askAI(prompt);
    n.note = (n.note.trim() ? n.note.trim() + "\n\n" : "") + `── AI 정리 (${hhmm()}) ──\n` + text;
    n.at = Date.now(); saveNow();
    const ta = $("note-" + mid); if (ta) ta.value = n.note;
    st.textContent = "메모에 추가했습니다."; syncBadges(); return true;
  } catch (e) {
    if (e.code === "nokey") { const how = await sendToClaude(prompt, `${m.org} 미팅 노트 정리`, btn); st.textContent = how === "share" ? "공유창에서 Claude 앱을 고르세요. 결과는 메모 칸에 붙여넣으면 됩니다." : how === "copy" ? "정리 요청문을 복사했습니다. Claude 앱에 붙여넣고 결과를 메모에 붙여넣으세요." : "보내기를 취소했습니다."; }
    else st.textContent = "정리하지 못했습니다: " + e.message;
    return false;
  } finally { btn.disabled = false; }
}

/* ---------- 하루 요약 ---------- */
const sumKey = (sid, d) => sid + "|" + d;
function dayMeetings(s, d) { return s.meetings.filter(m => m.date === d); }
function dayHasNotes(s, d) { return dayMeetings(s, d).some(m => { const n = note(s.id, m.id); return n.note.trim() || n.tr.trim(); }); }
function dayEnded(s, d) { const ms = dayMeetings(s, d); return new Date() >= kst(d, ms[ms.length - 1].end); }
function lastAt(s, d) { return Math.max(0, ...dayMeetings(s, d).map(m => note(s.id, m.id).at || 0)); }
function dayPrompt(s, d) {
  const body = dayMeetings(s, d).map(m => { const n = note(s.id, m.id); const t = n.note.trim() || n.tr.trim().slice(0, 8000); return t ? `### ${m.start}~${m.end} ${m.org} (${m.contact})\n${t}` : ""; }).filter(Boolean).join("\n\n");
  return `당신은 한국 셀사이드 애널리스트의 NDR 미팅 기록을 정리하는 보조입니다.
일정: ${s.title}. 날짜: ${d}(${dow(d)}). 아래는 기관투자자 미팅별 노트입니다.
기록에 없는 내용은 지어내지 말고, 불분명하면 "기록상 불명확"이라고 쓰세요.
다음 형식의 한국어 일반 텍스트로 간결하게 정리하세요(마크다운 기호 #, ** 쓰지 말 것):
■ 오늘 한 줄 요약
■ 기관별 핵심 (기관명: 주요 질문·관심사 / 반응 톤 긍정·중립·부정)
■ 공통적으로 나온 질문 Top 3
■ 회사·애널리스트 후속 확인 사항
■ 투자 포인트·주가 관련 시사점

${body}`;
}
function paintSum() {
  const s = cur(); if (!s) return; const d = ui.day;
  $("sumTitle").textContent = `${md(d)}(${dow(d)}) 요약`;
  const sm = S.sums[sumKey(s.id, d)], tx = $("sumText"), st = $("sumState");
  if (tx.dataset.busy === "1") return;
  if (sm && sm.text) {
    tx.hidden = false; tx.textContent = sm.text; $("sumRun").textContent = "다시 요약";
    st.textContent = lastAt(s, d) > sm.at ? " · 요약 이후 기록이 바뀌었습니다" : ` · ${new Date(sm.at).toLocaleString("ko-KR", { month: "numeric", day: "numeric", hour: "2-digit", minute: "2-digit" })} 작성`;
  } else {
    tx.hidden = true; $("sumRun").textContent = "하루 요약";
    st.textContent = !dayHasNotes(s, d) ? " · 아직 기록이 없습니다" : dayEnded(s, d) ? (settings().key ? " · 요약 대기" : " · '하루 요약'을 누르세요") : " · 마지막 미팅이 끝나면 자동으로 요약합니다";
  }
}
async function daySummary(s, d, btn) {
  if (!dayHasNotes(s, d)) { $("sumState").textContent = " · 요약할 기록이 없습니다"; return; }
  const prompt = dayPrompt(s, d), tx = $("sumText"), isCur = () => cur() === s && ui.day === d;
  if (isCur()) { tx.hidden = false; tx.dataset.busy = "1"; tx.textContent = "정리하는 중…"; $("sumState").textContent = " · 요약 작성 중"; }
  try {
    const text = await askAI(prompt);
    S.sums[sumKey(s.id, d)] = { text, at: Date.now() }; saveNow();
  } catch (e) {
    if (e.code === "nokey" && btn) { const how = await sendToClaude(prompt, `${md(d)} 하루 요약`, btn); if (isCur()) $("sumState").textContent = how === "share" ? " · 공유창에서 Claude 앱을 고르세요" : how === "copy" ? " · 요약 요청문을 복사했습니다. Claude 앱에 붙여넣으세요." : ""; }
    else if (isCur()) $("sumState").textContent = " · 요약하지 못했습니다: " + e.message;
    tx.dataset.busy = ""; if (!S.sums[sumKey(s.id, d)]) tx.hidden = true; return;
  }
  tx.dataset.busy = ""; if (isCur()) paintSum();
}
$("sumRun").onclick = e => { const s = cur(); if (s) daySummary(s, ui.day, e.currentTarget); };
$("sumCopy").onclick = e => {
  const s = cur(); if (!s) return; const d = ui.day, sm = S.sums[sumKey(s.id, d)];
  const txt = `${s.title} ${md(d)}(${dow(d)})\n\n` + (sm ? sm.text + "\n\n" : "") + "[미팅별 기록]\n" +
    dayMeetings(s, d).map(m => { const n = note(s.id, m.id); return `- ${m.start} ${m.org} (${m.contact})\n${n.note.trim() || n.tr.trim() || "(기록 없음)"}`; }).join("\n\n");
  copy(txt, e.currentTarget);
};
const autoTried = {};
async function maybeAuto() {
  if (!settings().key) return;
  for (const s of S.schedules) for (const d of dates(s)) {
    const k = sumKey(s.id, d);
    if (autoTried[k] || S.sums[k] || !dayEnded(s, d) || !dayHasNotes(s, d)) continue;
    if (Date.now() - kst(d, "00:00") > 3 * 86400000) continue; // 오래된 날짜는 자동 요약하지 않음
    autoTried[k] = true; await daySummary(s, d);
  }
}

/* ---------- 일정 관리 ---------- */
$("editText").onclick = () => {
  const s = cur(); const box = $("editBox"); box.hidden = !box.hidden; if (box.hidden) return;
  $("editArea").value = s.meetings.map(m => [m.date, m.start, m.end, m.org, m.contact, m.tel, m.loc, m.brokers.join("/")].join("\t")).join("\n");
};
$("saveEdit").onclick = () => {
  const s = cur(); const old = new Map(s.meetings.map(m => [m.date + m.start, m.id])); const out = []; const bad = [];
  $("editArea").value.split("\n").forEach((line, i) => {
    if (!line.trim()) return;
    const c = line.split("\t"); const date = SchedParser.normDate(c[0], new Date().getFullYear()), st = SchedParser.normTime(c[1]), en = SchedParser.normTime(c[2]);
    if (!date || !st) { bad.push(i + 1); return; }
    out.push({ id: old.get(date + st) || "m" + Math.random().toString(36).slice(2, 8), date, start: st, end: en || st, org: (c[3] || "").trim() || "(기관명 없음)", contact: (c[4] || "").trim(), tel: (c[5] || "").trim(), loc: (c[6] || "").trim(), brokers: (c[7] || "").split(/[\/,]/).map(x => x.trim()).filter(Boolean) });
  });
  if (!out.length) { msg("읽을 수 있는 줄이 없습니다."); return; }
  out.sort((a, b) => (a.date + a.start).localeCompare(b.date + b.start));
  s.meetings = out; saveNow(); $("editBox").hidden = true; renderSched();
  msg(`미팅 ${out.length}건으로 저장했습니다.` + (bad.length ? ` ${bad.join(", ")}번째 줄은 날짜·시간을 읽지 못해 뺐습니다.` : ""), !bad.length);
};
$("delSched").onclick = function () {
  if (this.dataset.arm !== "1") { this.dataset.arm = "1"; this.textContent = "정말 삭제 (메모 포함)"; setTimeout(() => { this.dataset.arm = ""; this.textContent = "이 일정 삭제"; }, 3500); return; }
  const s = cur(); if (R.on && R.sid === s.id) stopRec();
  S.schedules = S.schedules.filter(x => x.id !== s.id);
  Object.keys(S.notes).forEach(k => { if (k.startsWith(s.id + "|")) delete S.notes[k]; });
  Object.keys(S.sums).forEach(k => { if (k.startsWith(s.id + "|")) delete S.sums[k]; });
  S.cur = S.schedules[0]?.id || null; saveNow(); this.dataset.arm = ""; this.textContent = "이 일정 삭제";
  S.cur ? renderSched() : (resetImport(), show("import"));
};
$("exportBtn").onclick = () => {
  saveNow();
  const blob = new Blob([JSON.stringify({ app: "ndr-note", v: 1, exportedAt: new Date().toISOString(), data: S }, null, 1)], { type: "application/json" });
  const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = `ndr-note-backup-${todayStr().replace(/-/g, "")}.json`;
  document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
};
$("importBackup").onchange = async e => {
  const f = e.target.files[0]; e.target.value = ""; if (!f) return;
  try {
    const j = JSON.parse(await f.text()); const d = j.data; if (!d || !Array.isArray(d.schedules)) throw 0;
    d.schedules.forEach(s => { const i = S.schedules.findIndex(x => x.id === s.id); i >= 0 ? S.schedules[i] = s : S.schedules.push(s); });
    Object.entries(d.notes || {}).forEach(([k, v]) => { if (!S.notes[k] || (v.at || 0) > (S.notes[k].at || 0)) S.notes[k] = v; });
    Object.entries(d.sums || {}).forEach(([k, v]) => { if (!S.sums[k] || v.at > S.sums[k].at) S.sums[k] = v; });
    S.cur = S.cur || S.schedules[0]?.id; saveNow(); renderSched(); msg(`백업에서 일정 ${d.schedules.length}개를 불러왔습니다.`, true);
  } catch (err) { msg("백업 파일을 읽지 못했습니다."); }
};


/* ---------- API 없이: Claude 앱으로 보내기 · 정리 페이지 ---------- */
const DEFAULT_SUM_URL = "https://claude.ai/artifact/2eDDmyiEeoM7bumQCWbb7m";
function sumUrl() { return settings().sumUrl || DEFAULT_SUM_URL; }
async function sendToClaude(text, title, btn) {
  const mobile = /Android|iPhone|iPad/i.test(navigator.userAgent);
  if (mobile && navigator.share) {
    try { await navigator.share({ title, text }); return "share"; }
    catch (e) { if (e && e.name === "AbortError") return "cancel"; }
  }
  await copy(text, btn); return "copy";
}
function exportPack(s, ms) {
  const meetings = ms.map(m => { const n = note(s.id, m.id); return { date: m.date, start: m.start, end: m.end, org: m.org, contact: m.contact, tr: n.tr, note: n.note }; })
    .filter(x => x.tr.trim() || x.note.trim());
  return { text: "NDRNOTE:" + JSON.stringify({ app: "ndr-note", v: 1, title: s.title, meetings }), count: meetings.length };
}
function sendPack(s, ms, btn) {
  const { text, count } = exportPack(s, ms);
  const hint = $("sendHint");
  if (!count) { msg("보낼 녹취나 메모가 없습니다."); return; }
  const url = sumUrl(), label = btn.textContent;
  // 클립보드 복사와 새 창 열기를 같은 탭 동작 안에서 처리 (브라우저 차단 방지)
  const p = navigator.clipboard ? navigator.clipboard.writeText(text) : Promise.reject();
  window.open(url, "_blank", "noopener");
  p.then(() => { btn.textContent = `${count}건 복사됨`; msg(`미팅 ${count}건 기록을 복사했습니다. 열린 정리 페이지에 붙여넣고 '정리 시작'을 누르세요.`, true); })
   .catch(() => { msg("자동 복사가 막혀 있습니다. 아래 상자의 내용을 길게 눌러 복사하세요."); hint.hidden = false; hint.textContent = text; })
   .finally(() => setTimeout(() => btn.textContent = label, 2000));
}
$("sendDay").onclick = e => { const s = cur(); if (s) sendPack(s, dayMeetings(s, ui.day), e.currentTarget); };
$("sendAll").onclick = e => { const s = cur(); if (s) sendPack(s, s.meetings, e.currentTarget); };



/* ---------- 외부 녹음 앱 연결 ---------- */
const PLAY = id => encodeURIComponent("https://play.google.com/store/apps/details?id=" + id);
const REC_APPS = {
  browser: { name: "브라우저 받아쓰기" },
  clova: { name: "클로바노트", intent: `intent:#Intent;action=android.intent.action.MAIN;category=android.intent.category.LAUNCHER;package=com.naver.clova.minute;S.browser_fallback_url=${PLAY("com.naver.clova.minute")};end` },
  recorder: { name: "녹음 앱", intent: "intent:#Intent;action=android.provider.MediaStore.RECORD_SOUND;end" }
};
function recApp() { const a = settings().recApp; return REC_APPS[a] ? a : "browser"; }
const recName = m => `${m.date.slice(5).replace("-", "")} ${m.start.replace(":", "")} ${m.org}`;
function extStatus(n) {
  if (!n || !n.ext || n.ext.done) return "";
  return `${REC_APPS[n.ext.app]?.name || "녹음 앱"} 녹음 시작 ${new Date(n.ext.at).toLocaleTimeString("ko-KR", { hour: "2-digit", minute: "2-digit", hour12: false })} · 끝나면 녹취를 공유 → NDR 미팅노트로 보내세요`;
}
async function startExternal(sid, mid, app) {
  const s = S.schedules.find(x => x.id === sid), m = s.meetings.find(x => x.id === mid), n = note(sid, mid);
  if (R.on) stopRec();
  n.ext = { app, at: Date.now() }; n.at = Date.now(); saveNow();
  const name = recName(m); let copied = false;
  try { await navigator.clipboard.writeText(name); copied = true; } catch (e) { }
  const st = document.querySelector(`[data-rst="${mid}"]`);
  if (/Android/i.test(navigator.userAgent)) {
    if (st) st.textContent = `${REC_APPS[app].name}를 여는 중…` + (copied ? ` 녹음 제목 "${name}"이 복사돼 있습니다.` : "");
    location.href = REC_APPS[app].intent;
    setTimeout(() => { if (st) st.textContent = extStatus(n); }, 2500);
  } else {
    if (st) st.textContent = `이 기기에서는 ${REC_APPS[app].name}를 자동으로 열 수 없습니다. 직접 열어 녹음하세요.` + (copied ? ` 녹음 제목 "${name}"을 복사해 두었습니다.` : "") + " 끝나면 녹취를 복사해 '녹취 넣기'로 넣으세요.";
  }
}

/* ---------- 외부 녹취 넣기 (클로바노트·녹음 앱 공유 / 붙여넣기) ---------- */
let inboxSrc = "";
function findStamp(text) {
  // 녹음 시각 찾기: 2026.11.05 13:12 / 2026-11-05 오후 1:12 / 11월 5일 오후 1:12 / 11.05 13:12
  const t = String(text || "").slice(0, 600);
  const y0 = new Date().getFullYear();
  let m = t.match(/(20\d{2})\s*[.\-\/년]\s*(\d{1,2})\s*[.\-\/월]\s*(\d{1,2})\s*일?\.?\s*(?:\([^)]*\)\s*)?(오전|오후|AM|PM)?\s*(\d{1,2})\s*[:시]\s*(\d{2})?/i)
       || t.match(/()(\d{1,2})\s*월\s*(\d{1,2})\s*일\s*(?:\([^)]*\)\s*)?(오전|오후)?\s*(\d{1,2})\s*[:시]\s*(\d{2})?/);
  if (!m) return null;
  let h = +m[5]; const mi = +(m[6] || 0); const ap = (m[4] || "").toUpperCase();
  if ((ap === "오후" || ap === "PM") && h < 12) h += 12; if ((ap === "오전" || ap === "AM") && h === 12) h = 0;
  const y = m[1] ? +m[1] : y0;
  return { date: `${y}-${pad(+m[2])}-${pad(+m[3])}`, time: `${pad(h)}:${pad(mi)}` };
}
function guessMeeting(s, text) {
  const st = findStamp(text);
  if (st) {
    const t = kst(st.date, st.time);
    const hit = s.meetings.find(m => m.date === st.date && t >= new Date(kst(m.date, m.start) - 20 * 60000) && t <= new Date(+kst(m.date, m.end) + 10 * 60000));
    if (hit) return { m: hit, why: `녹음 시각 ${md(st.date)} ${st.time} 기준으로 골랐습니다.` };
  }
  const head = String(text || "").slice(0, 400);
  const byName = s.meetings.find(m => head.includes(recName(m))) || s.meetings.find(m => head.includes(m.date.slice(5).replace("-", "") + " " + m.start.replace(":", "")));
  if (byName) return { m: byName, why: "녹음 제목으로 골랐습니다." };
  const ext = s.meetings.map(m => ({ m, n: note(s.id, m.id) })).filter(x => x.n.ext && !x.n.ext.done && Date.now() - x.n.ext.at < 8 * 3600000).sort((a, b) => b.n.ext.at - a.n.ext.at)[0];
  if (ext) return { m: ext.m, why: `${REC_APPS[ext.n.ext.app]?.name || "녹음 앱"} 녹음을 시작한 미팅으로 골랐습니다.` };
  const now = new Date(), today = todayStr();
  const todays = s.meetings.filter(m => m.date === today);
  const live = todays.find(m => now >= kst(m.date, m.start) && now < kst(m.date, m.end));
  if (live) return { m: live, why: "지금 진행 중인 미팅으로 골랐습니다." };
  const past = todays.filter(m => kst(m.date, m.end) <= now).pop();
  if (past) return { m: past, why: "가장 최근에 끝난 미팅으로 골랐습니다. 다르면 바꿔주세요." };
  return { m: s.meetings.find(m => m.date === ui.day) || s.meetings[0], why: "녹음 시각을 찾지 못했습니다. 넣을 미팅을 골라주세요." };
}
function openInbox(text, src) {
  const s = cur(); if (!s) { msg("먼저 일정을 만들어 주세요."); return; }
  inboxSrc = src || "";
  $("inbox").hidden = false; $("inText").value = text || "";
  $("inSrc").textContent = src ? `"${src}"에서 받은 녹취입니다.` : "클로바노트·갤럭시 녹음 앱 등에서 녹취 텍스트를 복사해 오세요.";
  $("inMeeting").innerHTML = s.meetings.map(m => `<option value="${m.id}">${md(m.date)}(${dow(m.date)}) ${m.start} ${esc(m.org)}</option>`).join("");
  refreshGuess();
  $("inbox").scrollIntoView({ behavior: "smooth", block: "start" });
}
function refreshGuess() {
  const s = cur(); if (!s) return;
  const g = guessMeeting(s, inboxSrc + "\n" + $("inText").value);
  if (g.m) $("inMeeting").value = g.m.id;
  $("inGuess").textContent = $("inText").value.trim() ? g.why : "";
}
$("openInbox").onclick = () => openInbox("", "");
$("inText").addEventListener("input", () => { clearTimeout(refreshGuess.t); refreshGuess.t = setTimeout(refreshGuess, 400); });
$("inPaste").onclick = async () => {
  try { const t = await navigator.clipboard.readText(); if (t) { $("inText").value = t; refreshGuess(); } else msg("클립보드가 비어 있습니다."); }
  catch (e) { msg("클립보드를 읽을 수 없습니다. 칸을 길게 눌러 '붙여넣기' 하세요."); $("inText").focus(); }
};
$("inFile").onchange = async e => { const f = e.target.files[0]; e.target.value = ""; if (!f) return; inboxSrc = f.name; $("inText").value = await f.text(); refreshGuess(); };
$("inCancel").onclick = () => { $("inbox").hidden = true; };
async function inboxAdd(withAI) {
  const s = cur(), text = $("inText").value.trim(), mid = $("inMeeting").value;
  if (!text) { msg("넣을 녹취가 없습니다."); return; }
  const m = s.meetings.find(x => x.id === mid), n = note(s.id, mid);
  n.tr = (n.tr.trim() ? n.tr.replace(/\s+$/, "") + "\n\n" : "") + `── 외부 녹취${inboxSrc ? " · " + inboxSrc : ""} (${hhmm()} 추가) ──\n` + text + "\n";
  if (n.ext) n.ext.done = true;
  n.at = Date.now(); saveNow();
  $("inbox").hidden = true; ui.day = m.date; ui.broker = "전체"; ui.open.add(mid); renderSched();
  const el = document.querySelector(`.mtg[data-id="${mid}"]`); if (el) el.scrollIntoView({ behavior: "smooth", block: "start" });
  msg(`${m.org} 미팅 녹취에 넣었습니다.`, true);
  if (withAI) { const b = document.querySelector(`[data-ai="${mid}"]`); if (b) noteAI(s.id, mid, b); }
}
$("inAdd").onclick = () => inboxAdd(false);
$("inAddAI").onclick = () => inboxAdd(true);
async function checkShared() {
  if (!/[?&]shared=1/.test(location.search)) return;
  history.replaceState(null, "", location.pathname);
  try {
    const c = await caches.open("ndr-share"); const r = await c.match("shared"); if (!r) return;
    const d = await r.json(); await c.delete("shared");
    if (d.text) openInbox(d.text, d.title || "공유");
  } catch (e) { msg("공유받은 내용을 읽지 못했습니다. 복사해서 '녹취 넣기'로 붙여넣어 주세요."); }
}

/* ---------- 시작 ---------- */
if (S.schedules.length) { if (!cur()) S.cur = S.schedules[0].id; renderSched(); } else show("import");
checkShared();
setInterval(() => { if (!$("schedView").hidden) { refreshState(); if ($("sumText").dataset.busy !== "1") paintSum(); } maybeAuto(); }, 60000);
setTimeout(maybeAuto, 1500);
