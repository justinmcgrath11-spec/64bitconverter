(function(){
  'use strict';
  function byId(id){ return document.getElementById(id); }
  var els = {
    input: byId('input'), output: byId('output'), stats: byId('stats'), statsIn: byId('statsIn'), log: byId('log'),
    btnCopy: byId('btnCopy'), btnConvert: byId('btnConvert'), btnSave: byId('btnSave'),
    btnOpen: byId('btnOpen'), btnLoadDemo: byId('btnLoadDemo'), fileInput: byId('fileInput'),
    optAggressive: byId('optAggressive'), optVars: byId('optVars'),
    optWrap: byId('optWrap'), optComments: byId('optComments'), optAutoConvert: byId('optAutoConvert'),
    inputPanel: byId('inputPanel'),
    inputMirror: byId('inputMirror'), outputMirror: byId('outputMirror')
  };

  // --- Syntax highlight for apostrophe comments ---
  function escapeHTML(s){ return s.replace(/[&<>]/g,function(c){ return c==='&'?'&amp;':(c==='<'?'&lt;':'&gt;'); }); }
  function commentIndex(line){
    var inQ=false, i, ch;
    for(i=0;i<line.length;i++){
      ch=line.charAt(i);
      if(ch==='\"'){ inQ=!inQ; }
      else if(ch==='\''){ if(!inQ) return i; }
    }
    return -1;
  }
  function highlightVBA(text){
    var lines=text.split(/\r?\n/), out=[], i, line, idx, before, cmt;
    for(i=0;i<lines.length;i++){
      line=lines[i];
      idx=commentIndex(line);
      if(idx>=0){
        before=escapeHTML(line.slice(0,idx));
        cmt=escapeHTML(line.slice(idx));
        out.push(before + '<span class=\"cmt\">' + cmt + '</span>');
      }else{
        out.push(escapeHTML(line));
      }
    }
    return out.join('\n');
  }
  function bindMirror(textarea, mirror){
    function sync(){
      mirror.innerHTML = highlightVBA(textarea.value) + (textarea.value.slice(-1)==='\n' ? '\n' : '');
      mirror.style.height = textarea.scrollHeight + 'px';
    }
    textarea.addEventListener('input', sync);
    textarea.addEventListener('scroll', function(){ mirror.scrollTop = textarea.scrollTop; });
    sync();
    return sync;
  }

  var syncInput = bindMirror(els.input, els.inputMirror);
  var syncOutput = bindMirror(els.output, els.outputMirror);

  // --- Converter (same as safe v1) ---
  var pointerNameRx = /^(h|hwnd|hdc|hinst|hmenu|hicon|hcursor|hfont|hbrush|hbitmap|lp|lpsz|lpcstr|lpcwstr|lptstr|lpstr|lpwstr|p|ptr|pidl|ppv|ph|wparam|lparam|addr|address|window|handle|buffer|buf|pb|pv)/i;
  function stripInlineComment(line){ var inQ=false,out='',i,ch; for(i=0;i<line.length;i++){ ch=line.charAt(i); if(ch==='\"') inQ=!inQ; if(ch==='\''){ if(!inQ) break; } out+=ch; } return out; }
  function looksPointer(name){ return pointerNameRx.test(name) || /ptr|handle|wnd|addr/i.test(name); }
  function ensurePtrSafe(decl){ return decl.replace(/\bDeclare\b(?!\s+PtrSafe)/i,'Declare PtrSafe'); }
  function splitParams(s){ if(!s||!/\S/.test(s)) return []; var a=s.split(','),r=[],i; for(i=0;i<a.length;i++){ var t=a[i]; if(t&&/\S/.test(t)) r.push(t.replace(/^\s+|\s+$/g,'')); } return r; }

  function transformParam(p,chg){
    if(/As\s+Any\b/i.test(p)) return { txt:p, changed:false };
    var m=p.match(/^(\s*(?:Optional\s+)?(?:ByVal|ByRef)?\s*)([\w_]+)(\s*As\s+)([\w_]+)(.*)$/i);
    if(!m) return { txt:p, changed:false };
    var pre=m[1], name=m[2], asKw=m[3], typ=m[4], tail=m[5];
    if(/^Long$/i.test(typ) && looksPointer(name)){
      chg.params++; return { txt: pre+name+asKw+'LongPtr'+tail+(els.optComments&&els.optComments.checked?" ' [AutoConverted param Long->LongPtr]":""), changed:true };
    }
    return { txt:p, changed:false };
  }

  function transformReturn(line,fnName,chg){
    var m=line.match(/\)\s*As\s+(\w+)\s*$/i); if(!m) return { line:line, changed:false };
    var typ=m[1]; if(!/^Long$/i.test(typ)) return { line:line, changed:false };
    var looksLikeHandle=/ptr|window|handle|hwnd|getwindowlong/i.test(fnName);
    if(els.optAggressive&&els.optAggressive.checked&&looksLikeHandle){
      chg.returns++; return { line: line.replace(/\)\s*As\s+Long\s*$/i,") As LongPtr"+(els.optComments&&els.optComments.checked?" ' [AutoConverted return Long->LongPtr]":'')), changed:true };
    }
    return { line:line, changed:false };
  }

  function processDeclare(origLine,chg){
    var isFunction=/\bFunction\b/i.test(origLine);
    var paren=origLine.match(/\(([^)]*)\)/);
    var paramsTxt=paren?paren[1]:'';
    var before=paren?origLine.slice(0,paren.index):origLine;
    var after=paren?origLine.slice(paren.index+paren[0].length):'';

    var decl64=ensurePtrSafe(before)+'(';
    var parts=splitParams(paramsTxt);
    var newParams=[], any=false, i, r;
    for(i=0;i<parts.length;i++){ r=transformParam(parts[i],chg); newParams.push(r.txt); any=any||r.changed; }
    decl64+=newParams.join(', ')+')';

    if(isFunction){
      var fnName=(origLine.match(/\bFunction\s+(\w+)/i)||[])[1]||'';
      var res=transformReturn(decl64+after,fnName,chg);
      decl64=res.line; if(res.changed) any=true;
    } else {
      decl64=decl64+after;
    }

    if(!/\bPtrSafe\b/i.test(origLine)||any){ chg.declares++; }
    if(els.optComments&&els.optComments.checked&&decl64!==origLine&&!(/'\s*\[AutoConverted/.test(decl64))){ decl64+=" ' [AutoConverted declare]"; }
    if(els.optWrap&&els.optWrap.checked){
      var line32=origLine.replace(/\bPtrSafe\b/i,'').replace(/\s+\)/,')');
      return "#If VBA7 Then\n"+decl64+"\n#Else\n"+line32+"\n#End If";
    }
    return decl64;
  }

  function transformVariables(code,chg){
    if(!(els.optVars&&els.optVars.checked)) return code;
    var lines=code.split(/\r?\n/), i, line, body, segments, changed, rebuilt, s, seg, m, pre, name, asKw, tail;
    for(i=0;i<lines.length;i++){
      line=lines[i]; body=stripInlineComment(line);
      if(!/\b(Dim|Private|Public|Static)\b/i.test(body)) continue;
      if(!/\bAs\s+Long\b/i.test(body)) continue;
      segments=body.split(',');
      changed=false; rebuilt=[];
      for(s=0;s<segments.length;s++){
        seg=segments[s]; m=seg.match(/^(.*?)([\w_]+)(\s*As\s+)Long(\b.*)$/i);
        if(!m){ rebuilt.push(seg); continue; }
        pre=m[1]; name=m[2]; asKw=m[3]; tail=m[4];
        if(looksPointer(name)){ changed=true; chg.variables++; rebuilt.push(pre+name+asKw+'LongPtr'+tail+(els.optComments&&els.optComments.checked?" ' [AutoConverted var Long->LongPtr]":'')); }
        else{ rebuilt.push(seg); }
      }
      if(changed){
        var comment=line.indexOf("'")>=0?line.slice(line.indexOf("'")):"";
        lines[i]=rebuilt.join(',')+(comment&&!(els.optComments&&els.optComments.checked)?(" "+comment):"");
      }
    }
    return lines.join('\n');
  }

  function convert(code){
    var chg={declares:0,params:0,returns:0,variables:0};
    var logs=[];
    code=code.replace(/\s*_\r?\n\s*/g,' ');
    var lines=code.split(/\r?\n/), i, line, body, converted;
    for(i=0;i<lines.length;i++){
      line=lines[i]; body=stripInlineComment(line);
      if(/\bDeclare\b/i.test(body)){ converted=processDeclare(line,chg); logs.push('Updated Declare at line '+(i+1)); lines[i]=converted; }
    }
    var out=lines.join('\n'); out=transformVariables(out,chg);
    return { out:out, changes:chg, log:logs.join('\n') };
  }

  function run(){
    if(els.output){ els.output.readOnly=true; els.output.value=''; }
    var src=(els.input && typeof els.input.value==='string') ? els.input.value : '';
    var r=convert(src);
    if(els.output) els.output.value=r.out;
    if(els.stats) els.stats.textContent='Declares: '+r.changes.declares+' • Params->LongPtr: '+r.changes.params+' • Returns->LongPtr: '+r.changes.returns+' • Vars->LongPtr: '+r.changes.variables;
    if(els.log) els.log.textContent=r.log||'No changes yet.';
    if(els.statsIn) els.statsIn.textContent=(src.split(/\r?\n/).length)+' lines in input.';
    // refresh mirrors
    if (els.outputMirror) { els.outputMirror.innerHTML = highlightVBA(els.output.value); els.outputMirror.style.height = els.output.scrollHeight + 'px'; }
  }

  if(els.btnConvert) els.btnConvert.addEventListener('click', run);
  function onChange(){ run(); }
  var ids=['optAggressive','optVars','optWrap','optComments'];
  var i; for(i=0;i<ids.length;i++){ var o=byId(ids[i]); if(o) o.addEventListener('change', onChange); }

  if(els.btnCopy) els.btnCopy.addEventListener('click', function(){
    if(!els.output) return;
    els.output.readOnly=false; els.output.select();
    try{ document.execCommand('copy'); }catch(e){}
    els.output.readOnly=true;
    var t=this, old=t.textContent; t.textContent='Copied!'; setTimeout(function(){ t.textContent=old; }, 900);
  });

  if(els.btnSave) els.btnSave.addEventListener('click', function(){
    if(!els.output) return;
    var blob=new Blob([els.output.value],{type:'text/plain;charset=utf-8'});
    var a=document.createElement('a'); a.href=URL.createObjectURL(blob); a.download='converted_vba.bas'; a.click(); URL.revokeObjectURL(a.href);
  });

  if(els.btnOpen && els.fileInput){
    els.btnOpen.addEventListener('click', function(){ els.fileInput.click(); });
    els.fileInput.addEventListener('change', function(e){
      var f=e.target.files[0]; if(!f) return;
      var reader=new FileReader();
      reader.onload=function(){ els.input.value=reader.result; if(els.optAutoConvert && els.optAutoConvert.checked) run(); syncInput(); };
      reader.readAsText(f);
    });
  }

  function prevent(e){ e.preventDefault(); e.stopPropagation(); }
  if(els.input){ ['dragenter','dragover','dragleave','drop'].forEach(function(ev){ els.input.addEventListener(ev, prevent); }); }
  if(els.inputPanel){
    ['dragenter','dragover','dragleave','drop'].forEach(function(ev){ els.inputPanel.addEventListener(ev, prevent); });
    els.inputPanel.addEventListener('drop', function(e){
      var f=e.dataTransfer.files && e.dataTransfer.files[0]; if(!f) return;
      if(!/\.(bas|cls|txt)$/i.test(f.name)){ alert('Please drop a .bas, .cls, or .txt file'); return; }
      var reader=new FileReader();
      reader.onload=function(){ els.input.value=reader.result; if(els.optAutoConvert && els.optAutoConvert.checked) run(); syncInput(); };
      reader.readAsText(f);
    });
  }

  if(els.btnLoadDemo){
    els.btnLoadDemo.addEventListener('click', function(){
      var lines=[
        "' 32-bit demo macro using Win32 APIs",
        "' Convert this to 64-bit using the options below.",
        "",
        "Private Declare Function FindWindow Lib \"user32\" Alias \"FindWindowA\" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long",
        "Private Declare Function GetWindowLong Lib \"user32\" Alias \"GetWindowLongA\" (ByVal hWnd As Long, ByVal nIndex As Long) As Long",
        "Private Declare Sub CopyMemory Lib \"kernel32\" Alias \"RtlMoveMemory\" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)",
        "",
        "Public Sub Demo32()",
        "    Dim hWnd As Long",
        "    Dim ret As Long",
        "    ' Get Excel main window handle",
        "    hWnd = FindWindow(\"XLMAIN\", vbNullString)",
        "    If hWnd <> 0 Then",
        "        ' GWL_STYLE = -16",
        "        ret = GetWindowLong(hWnd, -16)",
        "    End If",
        "End Sub",
        ""
      ];
      els.input.value = lines.join('\n');
      if(els.optAutoConvert && els.optAutoConvert.checked) run();
      syncInput();
    });
  }

  if(els.statsIn){ els.statsIn.textContent='Paste VBA or drop a .bas/.cls, set options, click Convert.'; }
})();