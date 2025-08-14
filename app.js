// Environment guards (capture wallet/metamask errors so UI still loads)
const env = { warnings: [] };
const envNotes = document.getElementById('envNotes');
const envStatus = document.getElementById('envStatus');
function pushEnvWarning(msg){
  env.warnings.push(msg);
  if (envStatus) envStatus.classList.add('error');
  if (envNotes) envNotes.textContent = env.warnings.map((w,i)=>`[${i+1}] ${w}`).join('  ');
}
window.addEventListener('error', (e) => {
  const msg = (e && e.message) || '';
  if (/metamask|ethereum|wallet/i.test(msg)) {
    e.preventDefault && e.preventDefault();
    pushEnvWarning('External wallet error captured and ignored: ' + msg);
  }
});
window.addEventListener('unhandledrejection', (e) => {
  const reason = e && (e.reason && (e.reason.message || String(e.reason)));
  if (/metamask|ethereum|wallet/i.test(reason||'')){
    e.preventDefault && e.preventDefault();
    pushEnvWarning('External wallet rejection captured and ignored: ' + reason);
  }
});

// Elements
const els = {
  input: document.getElementById('input'),
  output: document.getElementById('output'),
  stats: document.getElementById('stats'),
  statsIn: document.getElementById('statsIn'),
  log: document.getElementById('log'),
  btnCopy: document.getElementById('btnCopy'),
  btnConvert: document.getElementById('btnConvert'),
  btnSave: document.getElementById('btnSave'),
  btnOpen: document.getElementById('btnOpen'),
  btnLoadDemo: document.getElementById('btnLoadDemo'),
  fileInput: document.getElementById('fileInput'),
  optAggressive: document.getElementById('optAggressive'),
  optVars: document.getElementById('optVars'),
  optWrap: document.getElementById('optWrap'),
  optComments: document.getElementById('optComments'),
  btnRunTests: document.getElementById('btnRunTests'),
  testLog: document.getElementById('testLog'),
  testSummary: document.getElementById('testSummary')
};

// Converter logic
const pointerNameRx = /^(h|hwnd|hdc|hinst|hmenu|hicon|hcursor|hfont|hbrush|hbitmap|lp|lpsz|lpcstr|lpcwstr|lptstr|lpstr|lpwstr|p|ptr|pidl|ppv|ph|wparam|lparam|addr|address|window|handle|buffer|buf|pb|pv)/i;

function stripInlineComment(line){
  let inQ = false, out = '';
  for (let i=0;i<line.length;i++){
    const ch = line[i];
    if (ch === '"') inQ = !inQ;
    if (ch === "'" && !inQ) break;
    out += ch;
  }
  return out;
}

function looksPointer(name){
  return pointerNameRx.test(name) || /ptr|handle|wnd|addr/i.test(name);
}

function ensurePtrSafe(decl){
  return decl.replace(/\bDeclare\b(?!\s+PtrSafe)/i, 'Declare PtrSafe');
}

function splitParams(paramStr){
  if (!paramStr.trim()) return [];
  return paramStr.split(',').map(s=>s.trim()).filter(Boolean);
}

function transformParam(p, opts, changes){
  const anyMatch = /As\s+Any\b/i.test(p);
  if (anyMatch) return { txt: p, changed: false };

  const m = p.match(/^(\s*(?:Optional\s+)?(?:ByVal|ByRef)?\s*)([\w_]+)(\s*As\s+)([\w_]+)(.*)$/i);
  if (!m) return { txt: p, changed: false };
  const [, pre, name, asKw, type, tail] = m;
  if (/^Long$/i.test(type) && looksPointer(name)){
    const newTxt = pre + name + asKw + 'LongPtr' + tail + (opts.comments? " ' [AutoConverted param Long→LongPtr]" : '');
    changes.params++;
    return { txt: newTxt, changed: true };
  }
  return { txt: p, changed: false };
}

function transformReturn(line, fnName, opts, changes){
  const m = line.match(/\)\s*As\s+(\w+)\s*$/i);
  if (!m) return { line, changed:false };
  const type = m[1];
  if (!/^Long$/i.test(type)) return { line, changed:false };
  const looksLikeHandle = /ptr|window|handle|hwnd|getwindowlong/i.test(fnName);
  if (opts.aggressive && looksLikeHandle){
    const newLine = line.replace(/\)\s*As\s+Long\s*$/i, ") As LongPtr" + (opts.comments? " ' [AutoConverted return Long→LongPtr]" : ''));
    changes.returns++;
    return { line: newLine, changed:true };
  }
  return { line, changed:false };
}

function processDeclare(origLine, opts, changes){
  const isFunction = /\bFunction\b/i.test(origLine);
  const paren = origLine.match(/\(([^)]*)\)/s);
  let paramsTxt = paren ? paren[1] : '';
  let before = paren ? origLine.slice(0, paren.index) : origLine;
  let after = paren ? origLine.slice(paren.index + paren[0].length) : '';

  let decl64 = ensurePtrSafe(before) + '(';

  const parts = splitParams(paramsTxt);
  const newParams = [];
  let anyParamChanged = false;
  for (const p of parts){
    const { txt, changed } = transformParam(p, opts, changes);
    newParams.push(txt);
    anyParamChanged = anyParamChanged || changed;
  }

  decl64 += newParams.join(', ') + ')';

  if (isFunction){
    const fnName = (origLine.match(/\bFunction\s+(\w+)/i)||[])[1] || '';
    const res = transformReturn(decl64 + after, fnName, opts, changes);
    decl64 = res.line;
    if (res.changed) anyParamChanged = true;
  } else {
    decl64 = decl64 + after;
  }

  if (!/\bPtrSafe\b/i.test(origLine) || anyParamChanged){
    changes.declares++;
  }

  if (opts.comments && decl64 !== origLine){
    if (!/\'\s*\[AutoConverted/.test(decl64)){
      decl64 += " ' [AutoConverted declare]";
    }
  }

  if (opts.wrap){
    const line32 = origLine.replace(/\bPtrSafe\b/i, '').replace(/\s+\)/, ')');
    return `#If VBA7 Then\n${decl64}\n#Else\n${line32}\n#End If`;
  }
  return decl64;
}

function transformVariables(code, opts, changes){
  if (!opts.vars) return code;
  const lines = code.split(/\r?\n/);
  for (let i=0;i<lines.length;i++){
    const line = lines[i];
    const body = stripInlineComment(line);
    if (!/\b(Dim|Private|Public|Static)\b/i.test(body)) continue;
    if (!/\bAs\s+Long\b/i.test(body)) continue;
    const segments = body.split(',');
    let changedAny = false;
    const rebuilt = segments.map(seg => {
      const m = seg.match(/^(.*?)([\w_]+)(\s*As\s+)Long(\b.*)$/i);
      if (!m) return seg;
      const [_, pre, name, asKw, tail] = m;
      if (looksPointer(name)){
        changedAny = true;
        changes.variables++;
        return pre + name + asKw + 'LongPtr' + tail + (opts.comments? " ' [AutoConverted var Long→LongPtr]" : '');
      }
      return seg;
    }).join(',');
    if (changedAny){
      const comment = line.includes("'") ? line.slice(line.indexOf("'")) : '';
      lines[i] = rebuilt + (comment && !opts.comments ? (' ' + comment) : '');
    }
  }
  return lines.join('\n');
}

function convert(code, opts){
  const changes = { declares:0, params:0, returns:0, variables:0 };
  const logLines = [];

  code = code.replace(/\s*_\r?\n\s*/g, ' ');

  const lines = code.split(/\r?\n/);
  for (let i=0; i<lines.length; i++){
    let line = lines[i];
    const body = stripInlineComment(line);
    if (/\bDeclare\b/i.test(body)){
      const converted = processDeclare(line, opts, changes);
      logLines.push('Updated Declare at line ' + (i+1));
      lines[i] = converted;
    }
  }

  let out = lines.join('\n');
  out = transformVariables(out, { vars: opts.vars, comments: opts.comments }, changes);

  return { out, changes, log: logLines.join('\n') };
}

// UI
function run(){
  const opts = {
    aggressive: els.optAggressive.checked,
    vars: els.optVars.checked,
    wrap: els.optWrap.checked,
    comments: els.optComments.checked,
  };
  const src = els.input.value || '';
  const { out, changes, log } = convert(src, opts);
  els.output.value = out;
  els.stats.textContent = `Declares updated: ${changes.declares} • Params Long→LongPtr: ${changes.params} • Return types changed: ${changes.returns} • Vars Long→LongPtr: ${changes.variables}`;
  els.log.textContent = log || 'No changes yet.';
  els.statsIn.textContent = `${src.split(/\r?\n/).length} lines in input.`;
}

if (els.btnConvert) els.btnConvert.addEventListener('click', run);
['optAggressive','optVars','optWrap','optComments'].forEach(id=>{
  const el = els[id];
  if (el) el.addEventListener('change', run);
});

if (els.btnCopy) els.btnCopy.addEventListener('click', async () => {
  els.output.select();
  document.execCommand('copy');
  els.btnCopy.textContent = 'Copied!';
  setTimeout(()=> els.btnCopy.textContent = 'Copy to Clipboard', 1200);
});

if (els.btnSave) els.btnSave.addEventListener('click', () => {
  const blob = new Blob([els.output.value], { type: 'text/plain;charset=utf-8' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'converted_vba.bas';
  a.click();
  URL.revokeObjectURL(a.href);
});

if (els.btnOpen && els.fileInput){
  els.btnOpen.addEventListener('click', ()=> els.fileInput.click());
  els.fileInput.addEventListener('change', async (e) => {
    const f = e.target.files[0];
    if (!f) return;
    const txt = await f.text();
    els.input.value = txt;
    run();
  });
}

if (els.btnLoadDemo){
  els.btnLoadDemo.addEventListener('click', () => {
    const demo = `Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)

Dim hWnd As Long, pBuf As Long, i As Long
`;
    els.input.value = demo;
    run();
  });
}

// Tests
function runUnitTests(){
  const results = [];
  function pass(name){ results.push({ name, ok:true }); }
  function fail(name, why){ results.push({ name, ok:false, why }); }
  function expectTrue(name, cond, context){ cond ? pass(name) : fail(name, context); }

  const base = { aggressive:false, vars:false, wrap:true, comments:true };

  (function(){
    const input = `Declare Function GetTickCount Lib "kernel32" () As Long`;
    const r = convert(input, base);
    expectTrue('T1: PtrSafe added', /Declare\s+PtrSafe\s+Function\s+GetTickCount/i.test(r.out), r.out);
    expectTrue('T1: Wrapped with VBA7', /#If\s+VBA7\s+Then[\s\S]*#Else[\s\S]*#End If/i.test(r.out), r.out);
  })();

  (function(){
    const input = `Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long`;
    const r = convert(input, base);
    expectTrue('T2: hWnd As LongPtr', /hWnd\s+As\s+LongPtr/i.test(r.out), r.out);
    expectTrue('T2: nIndex remains Long', /nIndex\s+As\s+Long\b/i.test(r.out), r.out);
  })();

  (function(){
    const input = `Dim hWnd As Long, i As Long`;
    const r = convert(input, { ...base, vars:true });
    expectTrue('T3: Var hWnd -> LongPtr', /hWnd\s+As\s+LongPtr/i.test(r.out), r.out);
    expectTrue('T3: Var i remains Long', /i\s+As\s+Long\b/i.test(r.out), r.out);
  })();

  (function(){
    const input = `Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
(ByVal lpClassName As String, _
 ByVal lpWindowName As String) As Long`;
    const r = convert(input, base);
    expectTrue('T4: PtrSafe on multiline declare', /Declare\s+PtrSafe\s+Function\s+FindWindow/i.test(r.out), r.out);
    expectTrue('T4: No raw line continuation sequences remain', !(/\s*_\r?\n\s*/.test(r.out)), r.out);
  })();

  (function(){
    const input = `Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Any, ByVal Source As Any, ByVal Length As Long)`;
    const r = convert(input, base);
    expectTrue('T5: Destination As Any kept', /Destination\s+As\s+Any/i.test(r.out), r.out);
    expectTrue('T5: Source As Any kept', /Source\s+As\s+Any/i.test(r.out), r.out);
  })();

  (function(){
    const before = env.warnings.length;
    const evt = new ErrorEvent('error', { message: 'Failed to connect to MetaMask' });
    window.dispatchEvent(evt);
    const after = env.warnings.length;
    expectTrue('T6: MetaMask error captured', after === before + 1, 'Warnings did not increment');
  })();

  const passed = results.filter(r=>r.ok).length;
  const failed = results.length - passed;
  if (els.testSummary) els.testSummary.innerHTML = failed === 0
    ? `<span class="tests-pass">All ${passed} tests passed ✔</span>`
    : `<span class="tests-fail">${failed} of ${results.length} tests failed ✖</span>`;

  if (els.testLog) els.testLog.textContent = results.map(r => r.ok ? `✅ ${r.name}` : `❌ ${r.name} — ${r.why || ''}`).join('\n');
}

if (els.btnRunTests) els.btnRunTests.addEventListener('click', runUnitTests);

// Initial tip
if (els.statsIn) els.statsIn.textContent = 'Paste VBA code on the left, click Convert.';
