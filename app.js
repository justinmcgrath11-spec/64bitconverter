const els = {
  inputCode: document.getElementById('inputCode'),
  outputCode: document.getElementById('outputCode'),
  convertBtn: document.getElementById('convertBtn'),
  optWrap: document.getElementById('optWrap'),
  optUpdateDeclares: document.getElementById('optUpdateDeclares')
};

els.convertBtn.addEventListener('click', () => {
  let code = els.inputCode.value;
  let converted = code;

  if (els.optUpdateDeclares.checked) {
    converted = converted.replace(/Long/g, 'LongPtr');
  }

  if (els.optWrap.checked) {
    converted = '#If VBA7 Then\n' + converted + '\n#Else\n' + code + '\n#End If';
  }

  els.outputCode.value = converted;
});
