var xlf = document.querySelector('#xlf');
var rABS = true; // true: readAsBinaryString ; false: readAsArrayBuffer
function handleFile(e) {
  console.log(rABS);
  var files = e.target.files, f = files[0];
  var reader = new FileReader();
  reader.onload = function(e) {
    var data = e.target.result;
    if(!rABS) data = new Uint8Array(data);
    var workbook = XLSX.read(data, {type: rABS ? 'binary' : 'array'});
    console.log(workbook);

    /* DO SOMETHING WITH workbook HERE */
  };
  if(rABS){
    reader.readAsBinaryString(f)
  }else {
    reader.readAsArrayBuffer(f)
  };
}
xlf.addEventListener('change', handleFile, false);
