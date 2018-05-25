var xlf = document.querySelector('#xlf');
if(xlf.addEventListener) xlf.addEventListener('change', handleFile, false);
function handleFile(e){
  var file  = e.target.files[0]
  var reader = new FileReader();
  reader.onload = function(e) {
    var data = e.target.result;
    if(!rABS) data = new Uint8Array(data);
    var workbook = XLSX.read(data, {type: rABS ? 'binary' : 'array'});

    /* DO SOMETHING WITH workbook HERE */
  };
  if(rABS) reader.readAsBinaryString(file); else reader.readAsArrayBuffer(file)
  cosole.log(workbook);
}
