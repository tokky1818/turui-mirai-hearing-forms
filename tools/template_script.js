/* ---------------- data (build.py が生成) ---------------- */
const FORM = __FORM_JSON__;

const STORAGE_KEY = "turui_hearing_" + FORM.key + "_v1";
const FORM_ID = "turui-" + FORM.key + "-v1";

/* ---------------- state ---------------- */
function localDate(){
  const d = new Date();
  return d.getFullYear() + "-" + String(d.getMonth()+1).padStart(2,"0") + "-" + String(d.getDate()).padStart(2,"0");
}
let state = { answers:{}, step:0, lastStep:0 };
try{
  const raw = localStorage.getItem(STORAGE_KEY);
  if(raw){ const parsed = JSON.parse(raw); if(parsed && typeof parsed==="object"){ state = Object.assign(state, parsed); } }
}catch(e){}
if(!state.answers) state.answers = {};
let hasResumed = Object.keys(state.answers).some(k=>k!=="R_DATE" && String(state.answers[k]).trim()!=="");
if(!state.answers.R_DATE){
  state.answers.R_DATE = localDate();
}

const BLOCKS = FORM.blocks;
const STEP_META = [{key:"cover", label:"はじめに"}]
  .concat(BLOCKS.map(b=>({key:b.id, label:b.railLabel})))
  .concat([{key:"docs", label:"提出書類チェック"}, {key:"review", label:"確認・保存"}]);
const STEP_DOCS = STEP_META.findIndex(s=>s.key==="docs");

const QUESTION_STEP = {};
BLOCKS.forEach((b,i)=>{ b.groups.forEach(g=>g.questions.forEach(q=>{ QUESTION_STEP[q.id] = i+1; })); });

let lastSavedAt = null;
function saveState(){
  try{
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
    lastSavedAt = new Date();
    updateSavedIndicator();
  }catch(e){}
}
let saveTimer=null;
function scheduleSave(){ clearTimeout(saveTimer); saveTimer=setTimeout(saveState, 250); }
function updateSavedIndicator(){
  const el = document.getElementById("savedIndicator");
  if(!el) return;
  if(!lastSavedAt){ el.textContent = ""; return; }
  const hh = String(lastSavedAt.getHours()).padStart(2,"0");
  const mm = String(lastSavedAt.getMinutes()).padStart(2,"0");
  el.textContent = `自動保存済み ${hh}:${mm}`;
}

/* ---------------- 設問の表示制御（Excelの「状態」列と同じ規則） ---------------- */
function norm(v){ return String(v==null?"":v).split("　").join(" ").trim(); }

/* 戻り値: "na"(対象外) / "pending"(先に選択) / "ok"(回答対象) */
function qGate(q){
  if(!q.dependsOn) return "ok";
  for(const c of q.dependsOn){
    const v = norm(state.answers[c.id]);
    if(v==="") return "pending";
    if(c.hideIf==="eq" && v===c.val) return "na";
    if(c.hideIf==="ne" && v!==c.val) return "na";
  }
  return "ok";
}
function isVisible(q){ return qGate(q)==="ok"; }
function qStatus(q){
  const g = qGate(q);
  if(g==="na") return "対象外";
  if(g==="pending") return "先に選択";
  return norm(state.answers[q.id])==="" ? "未入力" : "入力済";
}

function allQuestions(){
  const list=[];
  BLOCKS.forEach(b=>b.groups.forEach(g=>g.questions.forEach(q=>list.push(Object.assign({_blockId:b.id}, q)))));
  return list;
}
function allVisibleQuestions(){ return allQuestions().filter(isVisible); }

/* ブロック単位の状態: "na"=記入不要 / "pending"=先に別の設問へ回答 / "ok" */
function blockGate(b){
  let pending=false;
  for(const g of b.groups) for(const q of g.questions){
    const s = qGate(q);
    if(s==="ok") return "ok";
    if(s==="pending") pending=true;
  }
  return pending ? "pending" : "na";
}
function stepIsNA(i){
  const meta = STEP_META[i];
  const b = BLOCKS.find(x=>x.id===meta.key);
  return !!b && blockGate(b)==="na";
}

function computeProgress(){
  const qs = allVisibleQuestions();
  let total=qs.length, done=0;
  qs.forEach(q=>{ if(norm(state.answers[q.id])!=="") done++; });
  return {total, done};
}
function progressText(p){ return `${p.done}/${p.total} 回答済み`; }

/* ---------------- cross-device transfer ---------------- */
let pendingResumePayload = null;
let qrPanelOpen = false;
const RESUME_LINK_LIMIT = 1800;

function utf8ToBase64(str){
  const bytes = new TextEncoder().encode(str);
  let binary = "";
  bytes.forEach(b => binary += String.fromCharCode(b));
  return btoa(binary);
}
function base64ToUtf8(str){
  const binary = atob(str);
  const bytes = new Uint8Array(binary.length);
  for(let i=0;i<binary.length;i++) bytes[i] = binary.charCodeAt(i);
  return new TextDecoder().decode(bytes);
}
function buildTransferPayload(){
  return { formId: FORM_ID, savedAt: new Date().toISOString(), answers: state.answers };
}
function buildResumeLink(){
  const encoded = utf8ToBase64(JSON.stringify(buildTransferPayload()));
  return { url: location.origin + location.pathname + "#resume=" + encodeURIComponent(encoded), size: encoded.length };
}
function applyPayloadToState(payload){
  state.answers = payload.answers || {};
  state.step = 0;
  saveState();
  hasResumed = Object.keys(state.answers).length>0;
}
function clearResumeHash(){
  try{ history.replaceState(null, "", location.pathname + location.search); }catch(e){}
}
function isValidPayload(p){
  return !!(p && p.formId===FORM_ID && p.answers && typeof p.answers==="object" && !Array.isArray(p.answers));
}
function checkResumeHash(){
  if(location.hash && location.hash.indexOf("#resume=")===0){
    try{
      const encoded = decodeURIComponent(location.hash.slice(8));
      const payload = JSON.parse(base64ToUtf8(encoded));
      if(isValidPayload(payload)){ pendingResumePayload = payload; state.step = 0; }
    }catch(e){}
  }
}
function exportProgressFile(){
  trySaveFile(safeFileBase()+"_進捗.json", JSON.stringify(buildTransferPayload(), null, 2), "application/json");
}
function importProgressFile(file){
  const reader = new FileReader();
  reader.onload = ()=>{
    try{
      const payload = JSON.parse(reader.result);
      if(!isValidPayload(payload)){ showToast("このフォーム用の進捗ファイルではないようです"); return; }
      if(!confirm("読み込んだ内容で、この端末の入力内容を上書きします。よろしいですか？")) return;
      applyPayloadToState(payload);
      showToast("進捗ファイルを読み込みました");
      renderAll();
    }catch(e){
      showToast("ファイルの読み込みに失敗しました。正しい進捗ファイルかご確認ください");
    }
  };
  reader.readAsText(file, "utf-8");
}
function renderQrPanel(){
  const {size} = buildResumeLink();
  if(size > RESUME_LINK_LIMIT){
    return `<div class="qr-panel"><p class="qr-hint">現在の入力量が多く、QRコード／リンクでは送れません。上の「進捗をファイルに保存」をご利用ください。</p></div>`;
  }
  return `<div class="qr-panel">
    <div class="qr-canvas" id="qrCanvas"></div>
    <p class="qr-hint">QRコードを別端末のカメラで読み取るか、下のボタンでリンクをコピーしてください。入力内容が含まれるリンクのため、第三者と共有しないようご注意ください。</p>
    <div class="save-btn-row"><button class="btn" id="btnCopyResumeLink">リンクをコピーする</button></div>
  </div>`;
}

/* ---------------- rendering ---------------- */
const mainScroll = document.getElementById("mainScroll");
const railEl = document.getElementById("rail");
const trackerEl = document.getElementById("mobileTracker");
const footerEl = document.getElementById("footerBar");

function renderAll(){
  renderRail(); renderTracker(); renderStep(); renderFooter();
}

function stepClass(i){
  let cls = i===state.step ? "active" : (i<state.step ? "done" : "");
  if(stepIsNA(i)) cls += " na";
  return cls;
}
function renderRail(){
  let html = `<p class="rail-brand">${esc(FORM.brand)}<small>${esc(FORM.typeName)}</small></p>`;
  STEP_META.forEach((s,i)=>{
    html += `<div class="rail-step ${stepClass(i)}" data-step="${i}"><span class="rail-dot">${i}</span><span class="rail-label-wrap"><div class="rail-title">${esc(s.label)}</div></span></div>`;
  });
  railEl.innerHTML = html;
  railEl.querySelectorAll(".rail-step").forEach(el=>{
    el.addEventListener("click", ()=>{ goStep(parseInt(el.dataset.step,10)); });
  });
}

function renderTracker(){
  let html="";
  STEP_META.forEach((s,i)=>{
    html += `<div class="mt-pill ${stepClass(i)}" data-step="${i}">${esc(s.label)}</div>`;
  });
  trackerEl.innerHTML = html;
  trackerEl.querySelectorAll(".mt-pill").forEach(el=>{
    el.addEventListener("click", ()=>{ goStep(parseInt(el.dataset.step,10)); });
  });
  const active = trackerEl.querySelector(".mt-pill.active");
  if(active) active.scrollIntoView({inline:"center", block:"nearest"});
}

function fieldHtml(q){
  const val = state.answers[q.id] ?? "";
  const answered = String(val).trim()!=="";
  let inputHtml="";
  if(q.type==="select" && q.options.length<=4){
    inputHtml = `<div class="opt-group" role="radiogroup" aria-label="${escAttr(q.label)}">` +
      q.options.map(o=>`<button type="button" class="opt-chip ${o===val?"on":""}" role="radio" aria-checked="${o===val}" data-opt="${q.id}" data-val="${escAttr(o)}">${esc(o)}</button>`).join("") + `</div>`;
  } else if(q.type==="select"){
    inputHtml = `<select class="q-input" id="f_${q.id}" data-qid="${q.id}"><option value="">選択してください</option>` +
      q.options.map(o=>`<option value="${escAttr(o)}" ${o===val?"selected":""}>${esc(o)}</option>`).join("") + `</select>`;
  } else if(q.type==="textarea"){
    inputHtml = `<textarea class="q-input${q.short?" short":""}" id="f_${q.id}" data-qid="${q.id}" placeholder="ここに入力（分からなければ「わからない」）">${esc(val)}</textarea>`;
  } else {
    inputHtml = `<input class="q-input" type="text" id="f_${q.id}" data-qid="${q.id}" value="${escAttr(val)}" placeholder="ここに入力"${q.inputmode?` inputmode="${q.inputmode}"`:""}>`;
  }
  return `<div class="qcard ${answered?'answered':''}" id="card_${q.id}" data-card="${q.id}" ${isVisible(q)?"":"hidden"}>
    <div class="q-top"><span class="q-no">${esc(q.id)}</span></div>
    <label class="q-label" for="f_${q.id}">${esc(q.label)}</label>
    ${q.prompt?`<p class="q-prompt">${esc(q.prompt)}</p>`:''}
    ${q.example && q.type!=="select"?`<div class="q-example"><b>記入例：</b>${esc(q.example)}</div>`:''}
    ${inputHtml}
  </div>`;
}

function renderCover(){
  const c = FORM.cover;
  let html = `<div class="cover-wrap">
    <a class="back-link" href="index.html">← 類型の選択にもどる</a>
    <h1 class="cover-title serif">${esc(FORM.brand)}<br>事業計画ヒアリングフォーム</h1>
    <p class="cover-sub">${esc(FORM.typeName)}用　／　補助事業の事業計画書づくりに必要な情報を、一問一答形式でお伺いします。</p>`;
  if(pendingResumePayload){
    const savedAtStr = pendingResumePayload.savedAt ? new Date(pendingResumePayload.savedAt).toLocaleString("ja-JP") : "";
    html += `<div class="resume-note hash-resume-note">
      <span>別の端末からの引き継ぎデータが見つかりました${savedAtStr?'（'+esc(savedAtStr)+'時点）':''}。読み込みますか？現在この端末にある入力内容は上書きされます。</span>
      <span style="display:flex;gap:8px;flex-wrap:wrap;">
        <button class="btn primary btn-sm" id="btnApplyResumeHash">読み込む</button>
        <button class="btn btn-sm" id="btnDismissResumeHash">無視する</button>
      </span>
    </div>`;
  }
  if(hasResumed){
    html += `<div class="resume-note"><span>前回入力の続きから再開できます。</span><button id="btnReset">最初からやり直す</button></div>`;
  }
  html += `<div class="cover-card">
      <h3>■ ご回答にあたって（3つのポイント）</h3>
      <ol class="step-list">${c.points.map(t=>`<li>${esc(t)}</li>`).join("")}</ol>
      <button class="btn primary start-btn" id="btnStart">${hasResumed ? "続きから入力する" : "回答をはじめる"} →</button>
    </div>`;
  if(c.info && c.info.length){
    html += `<div class="cover-card"><h3>■ ご提出について</h3><ul class="info-list">` +
      c.info.map(r=>`<li><b>${esc(r[0])}</b><span>${esc(r[1])}</span></li>`).join("") + `</ul></div>`;
  }
  html += `<details class="cover-card"><summary>くわしい使い方</summary>
      <div class="usage-grid" style="margin-top:12px;">${c.usage.map(r=>`<div class="usage-item"><b>${esc(r[0])}</b><span>${esc(r[1])}</span></div>`).join("")}</div>
    </details>`;
  if(c.summary && c.summary.length){
    html += `<details class="cover-card"><summary>この補助金の概要（補助規則より）</summary><ul class="info-list" style="margin-top:12px;">` +
      c.summary.map(r=>`<li><b>${esc(r[0])}</b><span>${esc(r[1])}</span></li>`).join("") + `</ul>
      <p class="cover-caveat">${esc(c.summaryCaveat)}</p></details>`;
  }
  html += `<div class="cover-card">
      <h3>■ 別の端末で続きを入力する</h3>
      <p>この端末に保存されている入力内容を、他のスマートフォンやパソコンに引き継ぐことができます。</p>
      <div class="save-btn-row">
        <button class="btn" id="btnExportProgress">続きを保存（ファイル）</button>
        <button class="btn" id="btnImportProgress">保存した続きを読み込む</button>
      </div>
      <input type="file" id="importFileInput" accept="application/json,.json" style="display:none">
      <button class="btn" type="button" id="btnToggleQr" style="width:100%;">${qrPanelOpen ? "QRコード／リンクを閉じる" : "QRコード／リンクで送る（入力が少ないときのみ）"}</button>
      ${qrPanelOpen ? renderQrPanel() : ''}
    </div>
  </div>`;
  return html;
}

function naNoteText(b){
  return blockGate(b)==="pending"
    ? "このブロックの設問は、先に前のブロックの設問へ回答すると表示されます。"
    : "これまでの回答内容により、このブロックはご記入不要です。「次へ」で先へお進みください。";
}

function renderBlockContent(b){
  let html = `<div class="block-header">
    <p class="block-eyebrow">${esc(b.eyebrow)}</p>
    <h2 class="block-title">${esc(b.title)}</h2>
    ${b.note?`<div class="block-note">${esc(b.note)}</div>`:''}
  </div>`;
  const navGroups = b.groups.map((g,idx)=> g.heading ? {heading:g.heading, id:`grp_${b.id}_${idx}`} : null).filter(Boolean);
  if(navGroups.length > 1){
    html += `<div class="block-jumpnav">` + navGroups.map(ng=>`<button type="button" class="jump-chip" data-jump-group="${ng.id}">${esc(ng.heading)}</button>`).join('') + `</div>`;
  }
  const gate = blockGate(b);
  html += `<div class="na-note" id="naNote" ${gate==="ok"?"hidden":""}>${esc(naNoteText(b))}</div>`;
  b.groups.forEach((g,idx)=>{
    const anyVisible = g.questions.some(isVisible);
    html += `<div class="grp" id="grp_${b.id}_${idx}" ${anyVisible?"":"hidden"}>`;
    if(g.heading) html += `<div class="group-heading">${esc(g.heading)}</div>`;
    g.questions.forEach(q=>{ html += fieldHtml(q); });
    html += `</div>`;
  });
  return html;
}

function renderChecklistBlock(){
  let html = `<div class="block-header">
    <p class="block-eyebrow">提出書類</p>
    <h2 class="block-title">ご提出いただきたい資料</h2>
    <div class="block-note">準備の状況を選んでください。資料はメール・郵送・手渡しのいずれでもかまいません。スマホで撮影した写真でも大丈夫です。</div>
  </div>`;
  html += FORM.docs.map(d=>{
    const st = state.answers["ST_"+d.id] || FORM.docStates[FORM.docDefault];
    const note = state.answers["NT_"+d.id] || "";
    const chips = FORM.docStates.map(s=>{
      let cls="";
      if(s===st){
        if(s==="準備済み") cls="on-ready";
        else if(s==="準備中") cls="on-mid";
        else if(s==="該当なし") cls="on-none";
        else cls="on-todo";
      }
      const mark = s==="準備済み"?"☑ ":(s==="準備中"?"◐ ":(s==="まだ"?"☐ ":"－ "));
      return `<button class="doc-chip ${cls}" data-doc-state="${d.id}" data-val="${s}">${mark}${s}</button>`;
    }).join("");
    return `<div class="doc-row"><div class="doc-main"><div class="doc-label">${esc(d.no)}. ${esc(d.label)}</div><div class="doc-state">${chips}</div>
      <div class="doc-note"><input type="text" placeholder="備考・メモ（任意）" data-doc-note="${d.id}" value="${escAttr(note)}"></div>
    </div></div>`;
  }).join("");
  return html;
}

function renderReview(){
  const p = computeProgress();
  const missing = allVisibleQuestions().filter(q=>norm(state.answers[q.id])==="");
  let html = `<div class="block-header">
    <p class="block-eyebrow">最終ステップ</p>
    <h2 class="block-title">内容の確認・保存</h2>
    <p class="block-subtitle">ここまでの回答を確認し、ファイルに保存またはコピーして、担当の支援者へお送りください。</p>
  </div>
  <div class="review-stats">
    <div class="stat-tile"><div class="stat-num">${p.done} / ${p.total}</div><div class="stat-label">回答済みの設問数</div></div>
    <div class="stat-tile"><div class="stat-num">${missing.length}</div><div class="stat-label">未回答の設問数</div></div>
  </div>`;
  if(missing.length){
    html += `<div class="missing-box"><h4>あと ${missing.length} 問、未記入の設問があります</h4>
      <p style="margin:0 0 10px;font-size:12.5px;color:var(--text-muted);line-height:1.7;">当てはまらない・特にない場合は「なし」、分からない・未定の場合は「わからない」「未定」とご記入ください。項目をタップすると、その設問へ移動します。</p><ul class="missing-list">` +
      missing.slice(0,20).map(q=>`<li><button type="button" class="link-btn" data-jump="${q.id}">${esc(q.id)}：${esc(q.label)}</button></li>`).join("") +
      (missing.length>20?`<li>ほか ${missing.length-20} 件</li>`:"") + `</ul></div>`;
  } else if(p.total>0) {
    html += `<div class="done-box"><h4>✓ すべての設問に回答済みです</h4><p>この内容で保存・送付いただけます。提出書類の準備状況もあわせてご確認ください。</p></div>`;
  }
  const dest = (FORM.cover.info||[]).find(r=>r[0]==="返送先");
  html += `<div class="save-section">
    <h3>📝 提出の手順</h3>
    <ol class="step-list">
      <li>未記入の設問がないか確認します（上の一覧から戻って入力できます）。</li>
      <li>下の「テキストファイルを保存する」を押します。<br>保存できない場合は「全文をコピーする」を押してください。</li>
      <li>保存したファイル（またはコピーした文章）を、${dest?esc(dest[1]):"担当の支援者"}へメールやLINEなどでお送りください。${dest?"連絡先は事業者様へ個別にお知らせしています。":""}</li>
    </ol>
    <div class="save-btn-row">
      <button class="btn primary" id="btnSaveTxt">テキストファイルを保存する</button>
      <button class="btn" id="btnCopyTxt">全文をコピーする</button>
    </div>
    <textarea class="copy-area" id="copyArea" readonly aria-label="回答の全文"></textarea>
  </div>
  <details class="save-section">
    <summary>表計算ソフト用（CSV形式）で保存する　※担当の支援者から指示があった場合</summary>
    <p class="hint" style="margin-top:12px;">元のヒアリングシートの「回答一覧」と同じ列構成（項目番号・ブロック・項目・おたずね・回答・状態）で保存します。</p>
    <div class="save-btn-row"><button class="btn" id="btnSaveCsv">CSVファイルを保存する</button></div>
  </details>
  <p class="reset-link"><button type="button" class="link-btn" id="btnResetBottom">回答をすべて消去して最初からやり直す</button></p>`;
  return html;
}

function renderStep(){
  const meta = STEP_META[state.step];
  let html="";
  if(meta.key==="cover") html = renderCover();
  else if(meta.key==="review") html = renderReview();
  else if(meta.key==="docs") html = renderChecklistBlock();
  else {
    const block = BLOCKS.find(b=>b.id===meta.key);
    html = `<div>${renderBlockContent(block)}</div>`;
  }
  mainScroll.innerHTML = html;
  attachStepEvents();
  if(meta.key==="review"){
    document.getElementById("copyArea").value = buildReadableText();
  }
  mainScroll.scrollTo({top:0});
  window.scrollTo(0,0);
}

function renderFooter(){
  const p = computeProgress();
  const pct = p.total ? Math.round((p.done/p.total)*100) : 0;
  const isCover = state.step===0, isReview = state.step===STEP_META.length-1;
  footerEl.innerHTML = `
    <div class="footer-progress">
      <div class="footer-progress-text"><span class="footer-label">${esc(STEP_META[state.step].label)}</span><span id="progressText">${progressText(p)}</span></div>
      <div class="footer-bar-track"><div class="footer-bar-fill" style="width:${pct}%"></div></div>
      <div id="savedIndicator" class="footer-saved"></div>
    </div>
    ${isCover?'':'<button class="btn" id="btnPrev">← 戻る</button>'}
    ${isReview?'':'<button class="btn primary" id="btnNext">'+(isCover?'始める':'次へ')+' →</button>'}
  `;
  const prev = document.getElementById("btnPrev");
  const next = document.getElementById("btnNext");
  if(prev) prev.addEventListener("click", ()=>stepBy(-1));
  if(next) next.addEventListener("click", ()=>{ if(isCover){ startFromCover(); } else { stepBy(1); } });
  updateSavedIndicator();
}

/* 記入不要（対象外）のブロックは「次へ／戻る」で自動的に飛ばす */
function unansweredInCurrentBlock(){
  const b = BLOCKS.find(x=>x.id===STEP_META[state.step].key);
  if(!b) return 0;
  let n=0;
  b.groups.forEach(g=>g.questions.forEach(q=>{ if(isVisible(q) && norm(state.answers[q.id])==="") n++; }));
  return n;
}
function stepBy(dir){
  if(dir>0){
    const n = unansweredInCurrentBlock();
    if(n>0 && !confirm(`このブロックに未記入の設問が ${n} 問あります。\nこのまま次へ進みますか？\n（あとで「確認・保存」画面から戻って入力できます。分からない場合は「わからない」とご記入ください）`)) return;
  }
  let i = state.step + dir;
  while(i>0 && i<STEP_META.length-1 && stepIsNA(i)) i += dir;
  if(i<=0 && dir<0){ goStep(0); return; }
  goStep(i);
}
function startFromCover(){
  if(hasResumed && state.lastStep>0 && state.lastStep<STEP_META.length){ goStep(state.lastStep); return; }
  hasResumed = true;
  let i = 1;
  while(i<STEP_META.length-1 && stepIsNA(i)) i++;
  goStep(i);
}

function goStep(i){
  if(i<0||i>=STEP_META.length) return;
  state.step = i;
  if(i>0) state.lastStep = i;
  saveState();
  renderAll();
}

function attachStepEvents(){
  mainScroll.querySelectorAll("[data-qid]").forEach(el=>{
    el.addEventListener("input", onFieldInput);
    el.addEventListener("change", onFieldInput);
  });
  mainScroll.querySelectorAll("[data-opt]").forEach(el=>{
    el.addEventListener("click", ()=>{
      const id = el.dataset.opt;
      const same = state.answers[id]===el.dataset.val;
      setAnswer(id, same ? "" : el.dataset.val);
      const grp = el.parentElement;
      grp.querySelectorAll(".opt-chip").forEach(b=>{ const on = !same && b===el; b.classList.toggle("on", on); b.setAttribute("aria-checked", String(on)); });
    });
  });
  mainScroll.querySelectorAll("[data-doc-state]").forEach(el=>{
    el.addEventListener("click", ()=>{
      state.answers["ST_"+el.dataset.docState] = el.dataset.val;
      scheduleSave();
      const sc = mainScroll.scrollTop;
      renderStep(); renderFooter();
      mainScroll.scrollTop = sc;
    });
  });
  mainScroll.querySelectorAll("[data-doc-note]").forEach(el=>{
    el.addEventListener("input", ()=>{ state.answers["NT_"+el.dataset.docNote] = el.value; scheduleSave(); });
  });
  const btnStart = document.getElementById("btnStart");
  if(btnStart) btnStart.addEventListener("click", startFromCover);
  const btnReset = document.getElementById("btnReset");
  if(btnReset) btnReset.addEventListener("click", doReset);
  const btnResetBottom = document.getElementById("btnResetBottom");
  if(btnResetBottom) btnResetBottom.addEventListener("click", doReset);
  const btnApplyResumeHash = document.getElementById("btnApplyResumeHash");
  if(btnApplyResumeHash) btnApplyResumeHash.addEventListener("click", ()=>{
    applyPayloadToState(pendingResumePayload);
    pendingResumePayload = null;
    clearResumeHash();
    showToast("引き継ぎデータを読み込みました");
    renderAll();
  });
  const btnDismissResumeHash = document.getElementById("btnDismissResumeHash");
  if(btnDismissResumeHash) btnDismissResumeHash.addEventListener("click", ()=>{
    pendingResumePayload = null;
    clearResumeHash();
    renderStep();
  });
  const btnExportProgress = document.getElementById("btnExportProgress");
  if(btnExportProgress) btnExportProgress.addEventListener("click", exportProgressFile);
  const btnImportProgress = document.getElementById("btnImportProgress");
  const importFileInput = document.getElementById("importFileInput");
  if(btnImportProgress && importFileInput) btnImportProgress.addEventListener("click", ()=>importFileInput.click());
  if(importFileInput) importFileInput.addEventListener("change", (e)=>{
    const file = e.target.files[0];
    if(file) importProgressFile(file);
    e.target.value = "";
  });
  const btnToggleQr = document.getElementById("btnToggleQr");
  if(btnToggleQr) btnToggleQr.addEventListener("click", ()=>{ qrPanelOpen = !qrPanelOpen; renderStep(); });
  const qrCanvas = document.getElementById("qrCanvas");
  if(qrCanvas){
    if(window.QRCode){
      qrCanvas.innerHTML = "";
      new QRCode(qrCanvas, { text: buildResumeLink().url, width:180, height:180, correctLevel: QRCode.CorrectLevel.M });
    } else {
      qrCanvas.textContent = "QRコードを生成できませんでした。下のボタンでリンクをコピーしてください。";
    }
  }
  const btnCopyResumeLink = document.getElementById("btnCopyResumeLink");
  if(btnCopyResumeLink) btnCopyResumeLink.addEventListener("click", ()=>{
    const {url} = buildResumeLink();
    navigator.clipboard.writeText(url).then(()=>showToast("リンクをコピーしました")).catch(()=>showToast("コピーできませんでした"));
  });
  mainScroll.querySelectorAll("[data-jump]").forEach(el=>{
    el.addEventListener("click", ()=>{ jumpToQuestion(el.dataset.jump); });
  });
  const btnSaveCsv = document.getElementById("btnSaveCsv");
  if(btnSaveCsv) btnSaveCsv.addEventListener("click", saveCsv);
  const btnSaveTxt = document.getElementById("btnSaveTxt");
  if(btnSaveTxt) btnSaveTxt.addEventListener("click", saveTxt);
  const btnCopyTxt = document.getElementById("btnCopyTxt");
  if(btnCopyTxt) btnCopyTxt.addEventListener("click", copyTxt);
  mainScroll.querySelectorAll("[data-jump-group]").forEach(el=>{
    el.addEventListener("click", ()=>{
      const target = document.getElementById(el.dataset.jumpGroup);
      if(target) target.scrollIntoView({behavior:"smooth", block:"start"});
    });
  });
  mainScroll.querySelectorAll("textarea.q-input").forEach(el=>{
    autoGrow(el);
    el.addEventListener("input", ()=>autoGrow(el));
  });
}
function autoGrow(el){
  el.style.height = "auto";
  el.style.height = el.scrollHeight + "px";
}

/* 回答に応じて、設問カード・グループ・ブロックの表示を再計算する（入力欄は再描画しない） */
function refreshVisibility(){
  const meta = STEP_META[state.step];
  const b = BLOCKS.find(x=>x.id===meta.key);
  if(!b) return;
  b.groups.forEach((g,idx)=>{
    let any=false;
    g.questions.forEach(q=>{
      const card = document.getElementById("card_"+q.id);
      if(!card) return;
      const vis = isVisible(q);
      if(vis) any=true;
      if(card.hidden === vis){
        card.hidden = !vis;
        if(vis){ const ta = card.querySelector("textarea"); if(ta) autoGrow(ta); }
      }
    });
    const grp = document.getElementById(`grp_${b.id}_${idx}`);
    if(grp) grp.hidden = !any;
  });
  const naNote = document.getElementById("naNote");
  if(naNote){
    const gate = blockGate(b);
    naNote.hidden = gate==="ok";
    naNote.textContent = naNoteText(b);
  }
}

function onFieldInput(e){ setAnswer(e.target.dataset.qid, e.target.value); }
function setAnswer(id, value){
  state.answers[id] = value;
  const card = document.getElementById("card_"+id);
  if(card){ card.classList.toggle("answered", String(value).trim()!=="") ; }
  scheduleSave();
  refreshVisibility();
  renderFooterLite();
  refreshNav();
}
function refreshNav(){
  railEl.querySelectorAll(".rail-step").forEach(el=>{
    el.classList.toggle("na", stepIsNA(parseInt(el.dataset.step,10)));
  });
  trackerEl.querySelectorAll(".mt-pill").forEach(el=>{
    el.classList.toggle("na", stepIsNA(parseInt(el.dataset.step,10)));
  });
}
function renderFooterLite(){
  const p = computeProgress();
  const pct = p.total ? Math.round((p.done/p.total)*100) : 0;
  const txt = document.getElementById("progressText");
  const fill = footerEl.querySelector(".footer-bar-fill");
  if(txt) txt.textContent = progressText(p);
  if(fill) fill.style.width = pct+"%";
}

function jumpToQuestion(qid){
  const targetStep = QUESTION_STEP[qid];
  if(!targetStep) return;
  goStep(targetStep);
  setTimeout(()=>{
    const el = document.getElementById("card_"+qid);
    if(el) el.scrollIntoView({behavior:"smooth", block:"center"});
  },60);
}

function doReset(){
  if(!confirm("これまでの回答をすべて消去して、最初からやり直しますか？この操作は取り消せません。")) return;
  try{ localStorage.removeItem(STORAGE_KEY); }catch(e){}
  state = { answers:{}, step:0, lastStep:0 };
  state.answers.R_DATE = localDate();
  hasResumed=false;
  lastSavedAt = null;
  renderAll();
}

/* ---------------- export builders ---------------- */
function esc(s){ return String(s??"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;"); }
function escAttr(s){ return esc(s).replace(/"/g,"&quot;"); }

/* Excel「回答一覧（支援者・AI読み込み用）」と同じ構成: 項目番号 / ブロック / 項目 / おたずね / 回答 / 状態 */
function buildRows(){
  const rows=[];
  BLOCKS.forEach(b=>{
    b.groups.forEach(g=>g.questions.forEach(q=>{
      const st = qStatus(q);
      const shown = st==="入力済" || st==="未入力";
      rows.push({no:q.id, section:b.sheetLabel, label:q.label, prompt:q.prompt, answer: shown ? (state.answers[q.id]||"") : "", status:st});
    }));
  });
  FORM.docs.forEach(d=>{
    rows.push({no:"書類"+d.no, section:"提出書類チェック", label:d.label, prompt:"", answer:state.answers["ST_"+d.id]||FORM.docStates[FORM.docDefault], status:state.answers["NT_"+d.id]||"", isDoc:true});
  });
  return rows;
}

function csvEscape(v){
  v = String(v??"");
  if(/^[=+@]/.test(v) || /^-[^0-9]/.test(v)) v = "'" + v;
  if(/[",\r\n]/.test(v)) return '"' + v.replace(/"/g,'""') + '"';
  return v;
}
function buildCsv(){
  const rows = buildRows();
  let out = "﻿" + ["項目番号","ブロック","項目","おたずね","回答","状態（提出書類は備考）"].map(csvEscape).join(",") + "\r\n";
  out += ["","記入日","","",state.answers.R_DATE||"",""].map(csvEscape).join(",") + "\r\n";
  rows.forEach(r=>{
    out += [r.no, r.section, r.label, r.prompt, r.answer, r.status].map(csvEscape).join(",") + "\r\n";
  });
  return out;
}

function buildReadableText(){
  const rows = buildRows();
  const NL = String.fromCharCode(10);
  let out = `【${FORM.brand}】事業計画ヒアリングフォーム（${FORM.typeName}） 回答控え` + NL;
  out += "事業者名：" + (state.answers["00-01"]||"") + "　代表者：" + (state.answers["00-02"]||"") + "　記入日：" + (state.answers.R_DATE||"") + NL;
  out += "生成日時：" + new Date().toLocaleString("ja-JP") + NL;
  out += "========================================" + NL;
  let curSection="";
  rows.forEach(r=>{
    if(r.status==="対象外" || r.status==="先に選択") return;
    if(r.section!==curSection){ curSection=r.section; out += NL + "■ "+curSection + NL; }
    if(r.isDoc){
      out += "・" + r.label + NL + "  → " + r.answer + (r.status?"（"+r.status+"）":"") + NL;
      return;
    }
    out += "・" + r.no + " " + r.label + (r.status==="未入力"?"【未回答】":"") + NL;
    out += "  → " + (r.answer || "（未回答）") + NL;
  });
  return out;
}

function safeFileBase(){
  const biz = String(state.answers["00-01"]||"事業者").replace(/[/:*?"<>|]/g,"").replace(/\s+/g," ").split(String.fromCharCode(92)).join("").trim().slice(0,30) || "事業者";
  const date = (state.answers.R_DATE||"").replace(/-/g,"");
  return `ヒアリング回答_${FORM.short}_${biz}_${date}`;
}

function showToast(msg){
  const t = document.getElementById("toast");
  t.textContent = msg;
  t.classList.add("show");
  clearTimeout(showToast._t);
  showToast._t = setTimeout(()=>t.classList.remove("show"), 2600);
}

function trySaveFile(filename, content, mime){
  try{
    const blob = new Blob([content], {type: mime + ";charset=utf-8"});
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = filename;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    setTimeout(()=>URL.revokeObjectURL(url), 4000);
    showToast("保存しました：" + filename);
  }catch(e){
    const ca = document.getElementById("copyArea");
    if(ca){ ca.value = content; ca.scrollIntoView({behavior:"smooth", block:"center"}); }
    showToast("自動保存に失敗したため、下の欄からコピーしてご利用ください");
  }
}
function saveCsv(){ trySaveFile(safeFileBase()+".csv", buildCsv(), "text/csv"); }
function saveTxt(){ trySaveFile(safeFileBase()+".txt", buildReadableText(), "text/plain"); }
async function copyTxt(){
  const text = buildReadableText();
  const ca = document.getElementById("copyArea");
  if(ca) ca.value = text;
  try{
    await navigator.clipboard.writeText(text);
    showToast("コピーしました");
  }catch(e){
    if(ca){ ca.focus(); ca.select(); try{ document.execCommand("copy"); showToast("コピーしました"); }catch(e2){ showToast("コピーできませんでした。テキストを選択してコピーしてください"); } }
  }
}

/* ---------------- init ---------------- */
window.addEventListener("pagehide", saveState);
document.addEventListener("visibilitychange", ()=>{ if(document.visibilityState==="hidden") saveState(); });
checkResumeHash();
if(state.step<0 || state.step>=STEP_META.length) state.step = 0;
renderAll();
