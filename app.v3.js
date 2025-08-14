// v3: robust loader + cache-busted file name. Hide JS warning once running.
(function(){
  var warn = document.getElementById('jsWarning');
  if (warn) warn.className = 'js-ok';

  function $id(id){ return document.getElementById(id); }
  const els = {
    input: $id('input'), output: $id('output'), stats: $id('stats'), statsIn: $id('statsIn'), log: $id('log'),
    btnCopy: $id('btnCopy'), btnConvert: $id('btnConvert'), btnSave: $id('btnSave'), btnOpen: $id('btnOpen'),
    btnLoadDemo: $id('btnLoadDemo'), fileInput: $id('fileInput'), optAggressive: $id('optAggressive'),
    optVars: $id('optVars'), optWrap: $id('optWrap'), optComments: $id('optComments'), optAutoConvert: $id('optAutoConvert'),
    btnRunTests: $id('btnRunTests'), testLog: $id('testLog'), testSummary: $id('testSummary'), inputPanel: $id('inputPanel')
  };

  // Converter logic (same as v2)
  const pointerNameRx = /^(h|hwnd|hdc|hinst|hmenu|hicon|hcursor|hfont|hbrush|hbitmap|lp|lpsz|lpcstr|lpcwstr|lptstr|lpstr|lpwstr|p|ptr|pidl|ppv|ph|wparam|lparam|addr|address|window|handle|buffer|buf|pb|pv)/i;
  function stripInlineComment(line){ let inQ=false,out=''; for(let i=0;i<line.length;i++){ const ch=line[i]; if(ch==='\"') inQ=!inQ; if(ch===\"'\"&&!inQ) break; out+=ch; } return out; }
  function looksPointer(name){ return pointerNameRx.test(name) || /ptr|handle|wnd|addr/i.test(name); }
  function ensurePtrSafe(decl){ return decl.replace(/\\bDeclare\\b(?!\\s+PtrSafe)/i, 'Declare PtrSafe'); }
  function splitParams(paramStr){ if(!paramStr.trim()) return []; return paramStr.split(',').map(s=>s.trim()).filter(Boolean); }
  function transformParam(p, opts, changes){
    if (/As\\s+Any\\b/i.test(p)) return { txt: p, changed: false };
    const m = p.match(/^(\\s*(?:Optional\\s+)?(?:ByVal|ByRef)?\\s*)([\\w_]+)(\\s*As\\s+)([\\w_]+)(.*)$/i); if(!m) return { txt:p, changed:false };
    const [,pre,name,asKw,type,tail]=m;
    if (/^Long$/i.test(type) && looksPointer(name)){ changes.params++; return { txt: pre+name+asKw+'LongPtr'+tail+(els.optComments.checked?\" ' [AutoConverted param Long→LongPtr]\":''), changed:true }; }
    return { txt:p, changed:false };
  }
  function transformReturn(line, fnName, opts, changes){
    const m = line.match(/\\)\\s*As\\s+(\\w+)\\s*$/i); if(!m) return { line, changed:false };
    const type = m[1]; if(!/^Long$/i.test(type)) return { line, changed:false };
    if (els.optAggressive.checked && /ptr|window|handle|hwnd|getwindowlong/i.test(fnName)){
      changes.returns++; return { line: line.replace(/\\)\\s*As\\s+Long\\s*$/i, \") As LongPtr\" + (els.optComments.checked?\" ' [AutoConverted return Long→LongPtr]\":'')), changed:true };
    }
    return { line, changed:false };
  }
  function processDeclare(origLine, opts, changes){
    const isFunction=/\\bFunction\\b/i.test(origLine); const paren=origLine.match(/\\(([^)]*)\\)/s);
    let paramsTxt=paren?paren[1]:''; let before=paren?origLine.slice(0,paren.index):origLine; let after=paren?origLine.slice(paren.index+paren[0].length):'';
    let decl64=ensurePtrSafe(before)+'(';
    const parts=splitParams(paramsTxt); const newParams=[]; let anyParamChanged=false;
    for(const p of parts){ const {txt,changed}=transformParam(p, {}, changes); newParams.push(txt); anyParamChanged = anyParamChanged || changed; }
    decl64+=newParams.join(', ')+')';
    if(isFunction){ const fnName=(origLine.match(/\\bFunction\\s+(\\w+)/i)||[])[1]||''; const res=transformReturn(decl64+after, fnName, {}, changes); decl64=res.line; if(res.changed) anyParamChanged=true; } else { decl64=decl64+after; }
    if(!/\\bPtrSafe\\b/i.test(origLine) || anyParamChanged){ changes.declares++; }
    if(els.optComments.checked && decl64 !== origLine && !(/\\'\\s*\\[AutoConverted/.test(decl64))){ decl64 += \" ' [AutoConverted declare]\"; }
    if(els.optWrap.checked){ const line32 = origLine.replace(/\\bPtrSafe\\b/i,'').replace(/\\s+\\)/,')'); return `#If VBA7 Then\\n${decl64}\\n#Else\\n${line32}\\n#End If`; }
    return decl64;
  }
  function transformVariables(code, opts, changes){
    if(!els.optVars.checked) return code; const lines=code.split(/\\r?\\n/);
    for(let i=0;i<lines.length;i++){ const line=lines[i]; const body=stripInlineComment(line);
      if(!/\\b(Dim|Private|Public|Static)\\b/i.test(body)) continue; if(!/\\bAs\\s+Long\\b/i.test(body)) continue;
      const segments=body.split(','); let changedAny=false;
      const rebuilt=segments.map(seg=>{ const m=seg.match(/^(.*?)([\\w_]+)(\\s*As\\s+)Long(\\b.*)$/i); if(!m)return seg; const [_,pre,name,asKw,tail]=m;
        if(looksPointer(name)){ changedAny=true; changes.variables++; return pre+name+asKw+'LongPtr'+tail+(els.optComments.checked?\" ' [AutoConverted var Long→LongPtr]\":''); } return seg; }).join(',');
      if(changedAny){ const comment=line.includes(\"'\")?line.slice(line.indexOf(\"'\")):''; lines[i]=rebuilt + (comment && !els.optComments.checked ? (' '+comment):''); }
    }
    return lines.join('\\n');
  }
  function convert(code){
    const changes={declares:0,params:0,returns:0,variables:0}; const logLines=[];
    code = code.replace(/\\s*_\\r?\\n\\s*/g,' ');
    const lines=code.split(/\\r?\\n/);
    for(let i=0;i<lines.length;i++){ let line=lines[i]; const body=stripInlineComment(line); if(/\\bDeclare\\b/i.test(body)){ const converted=processDeclare(line, {}, changes); logLines.push('Updated Declare at line '+(i+1)); lines[i]=converted; } }
    let out=lines.join('\\n'); out=transformVariables(out, {}, changes);
    return { out, changes, log: logLines.join('\\n') };
  }
  function run(){
    els.output.readOnly = true; els.output.value='';
    const src = els.input.value || ''; const { out, changes, log } = convert(src);
    els.output.value = out; if(els.stats) els.stats.textContent = `Declares: ${changes.declares} • Params→LongPtr: ${changes.params} • Returns→LongPtr: ${changes.returns} • Vars→LongPtr: ${changes.variables}`;
    if(els.log) els.log.textContent = log || 'No changes yet.'; if(els.statsIn) els.statsIn.textContent = `${src.split(/\\r?\\n/).length} lines in input.`;
  }

  // Buttons
  if (els.btnConvert) els.btnConvert.addEventListener('click', run);
  if (els.btnCopy) els.btnCopy.addEventListener('click', function(){
    els.output.readOnly=false; els.output.select(); document.execCommand('copy'); els.output.readOnly=true;
    var t=this; var old=t.textContent; t.textContent='Copied!'; setTimeout(()=>t.textContent=old, 900);
  });
  if (els.btnSave) els.btnSave.addEventListener('click', function(){
    const blob=new Blob([els.output.value],{type:'text/plain;charset=utf-8'}); const a=document.createElement('a'); a.href=URL.createObjectURL(blob); a.download='converted_vba.bas'; a.click(); URL.revokeObjectURL(a.href);
  });
  if (els.btnOpen && els.fileInput){
    els.btnOpen.addEventListener('click', ()=> els.fileInput.click());
    els.fileInput.addEventListener('change', async (e)=>{ const f=e.target.files[0]; if(!f) return; const txt = await f.text(); els.input.value=txt; if(els.optAutoConvert && els.optAutoConvert.checked) run(); });
  }

  // Drag & drop
  function prevent(e){ e.preventDefault(); e.stopPropagation(); }
  ['dragenter','dragover','dragleave','drop'].forEach(evt=>{ if(els.input) els.input.addEventListener(evt, prevent); if(els.inputPanel) els.inputPanel.addEventListener(evt, prevent); });
  if (els.inputPanel) els.inputPanel.addEventListener('drop', async (e)=>{ const f=e.dataTransfer.files && e.dataTransfer.files[0]; if(!f) return; if(!/\\.(bas|cls|txt)$/i.test(f.name)){ alert('Please drop a .bas, .cls, or .txt file'); return; } const txt=await f.text(); els.input.value=txt; if(els.optAutoConvert && els.optAutoConvert.checked) run(); });

  // Demo macro (32-bit)
  if (els.btnLoadDemo){
    els.btnLoadDemo.addEventListener('click', function(){
      const demo = `' 32-bit demo macro using Win32 APIs\n' Convert this to 64-bit using the options below.\n\nPrivate Declare Function FindWindow Lib \"user32\" Alias \"FindWindowA\" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long\nPrivate Declare Function GetWindowLong Lib \"user32\" Alias \"GetWindowLongA\" (ByVal hWnd As Long, ByVal nIndex As Long) As Long\nPrivate Declare Sub CopyMemory Lib \"kernel32\" Alias \"RtlMoveMemory\" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)\n\nPublic Sub Demo32()\n    Dim hWnd As Long\n    Dim ret As Long\n    ' Get Excel main window handle\n    hWnd = FindWindow(\"XLMAIN\", vbNullString)\n    If hWnd <> 0 Then\n        ' GWL_STYLE = -16\n        ret = GetWindowLong(hWnd, -16)\n    End If\nEnd Sub\n`;
      els.input.value = demo; if(els.optAutoConvert && els.optAutoConvert.checked) run();
    });
  }

  // Unit tests
  if (els.btnRunTests){
    els.btnRunTests.addEventListener('click', function(){
      const results=[]; function pass(n){results.push({n,ok:true})} function fail(n,w){results.push({n,ok:false,w})} function expect(n,c,ctx){ c?pass(n):fail(n,ctx) }
      const base = { };
      (function(){ const input=`Declare Function GetTickCount Lib "kernel32" () As Long`; const r=convert(input); expect('T1 PtrSafe', /Declare\\s+PtrSafe\\s+Function\\s+GetTickCount/i.test(r.out), r.out); expect('T1 Wrap', /#If\\s+VBA7\\s+Then[\\s\\S]*#Else[\\s\\S]*#End If/i.test(r.out), r.out); })();
      (function(){ const input=`Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long`; const r=convert(input); expect('T2 hWnd LongPtr', /hWnd\\s+As\\s+LongPtr/i.test(r.out), r.out); expect('T2 nIndex Long', /nIndex\\s+As\\s+Long\\b/i.test(r.out), r.out); })();
      (function(){ const input=`Dim hWnd As Long, i As Long`; const r=convert(input); expect('T3 Var hWnd LongPtr', /hWnd\\s+As\\s+LongPtr/i.test(r.out), r.out); expect('T3 i Long', /i\\s+As\\s+Long\\b/i.test(r.out), r.out); })();
      (function(){ const input=`Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _\n(ByVal lpClassName As String, _\n ByVal lpWindowName As String) As Long`; const r=convert(input); expect('T4 multiline PtrSafe', /Declare\\s+PtrSafe\\s+Function\\s+FindWindow/i.test(r.out), r.out); expect('T4 no continuations', !(/\\s*_\\r?\\n\\s*/.test(r.out)), r.out); })();
      (function(){ const input=`Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Any, ByVal Source As Any, ByVal Length As Long)`; const r=convert(input); expect('T5 Any kept', /Destination\\s+As\\s+Any/i.test(r.out), r.out); expect('T5 Any kept2', /Source\\s+As\\s+Any/i.test(r.out), r.out); })();
      const passed=results.filter(r=>r.ok).length,total=results.length,failed=total-passed; if(els.testSummary) els.testSummary.innerHTML = failed? `<span class='err'>${failed} of ${total} tests failed ✖</span>` : `<span class='ok'>All ${passed} tests passed ✔</span>`; if(els.testLog) els.testLog.textContent = results.map(r=> r.ok?`✅ ${r.n}`:`❌ ${r.n} — ${r.w||''}`).join('\\n');
    });
  }

  if (els.statsIn) els.statsIn.textContent = 'Paste VBA or drop a .bas/.cls, set options, click Convert.';
})();