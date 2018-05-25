const nodemailer = require('nodemailer');
var XLSX = require('xlsx')
var xlf = document.querySelector('#xlf');
var rABS = true; // true: readAsBinaryString ; false: readAsArrayBuffer
var loading = {}
var workbook;
var workbookJson;



function handleFile(e) {
  var files = e.target.files, f = files[0];
  var reader = new FileReader();
  reader.onload = function(e) {
    var data = e.target.result;
    if(!rABS) data = new Uint8Array(data);
    workbook = XLSX.read(data, {type: rABS ? 'binary' : 'array'});
    var workbookJson = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
    var html1 = ''
    workbookJson.map(data => {
      html1 += `<tr><td>${Object.values(data).join('</td><td>')}</td></tr>`;
    })
    var html = `<table border="1" style="border-collapse:collapse">
        <tr><td>${Object.keys(workbookJson[0]).join(`</td><td>`)}</td></tr>
        ${html1}
      </table>`;
    document.querySelector('#excel').innerHTML = html;

    /* DO SOMETHING WITH workbook HERE */
  };
  if(rABS){
    reader.readAsBinaryString(f)
  }else {
    reader.readAsArrayBuffer(f)
  };
  // return
}

function submit(e) {
  var email = document.querySelector('#email').value;
  var password = document.querySelector('#password').value;
  console.log(email, password);
  transfer(workbook, email, password)
}
xlf.addEventListener('change', handleFile, false);

function transfer(workbook, email, password) {
  // 读取excel文件
  // var workbook = XLSX.readFile(path);
  // 获取第一个sheet名
  var first_sheet_name = workbook.SheetNames[0];
  var address_of_cell = 'A1';
  // 获取第一个sheet
  var worksheet = workbook.Sheets[first_sheet_name];
  // console.log(worksheet);
  var range = XLSX.utils.decode_range(worksheet['!ref']);
  console.log(range);
  // 获取第一个表格头部
  var table_head = [];
  for (var j = range.s.c; j <= range.e.c; ++j) {
    var cell_address = {
      c: j,
      r: 0
    };
    var cell_ref = XLSX.utils.encode_cell(cell_address);
    table_head.push(worksheet[cell_ref].v)
  }

  for (var i = range.s.r + 1; i <= range.e.r; ++i) {
    var stemp = []
    // 取最后一列邮箱地址
    var addressCell = worksheet[XLSX.utils.encode_cell({
      c: range.e.c,
      r: i
    })];
    var address = addressCell.v;

    // 分别获取每一行的值
    for (var j = range.s.c; j <= range.e.c; ++j) {
      var cell_address = {
        c: j,
        r: i
      };
      var cell_ref = XLSX.utils.encode_cell(cell_address);
      stemp.push(worksheet[cell_ref].v)
    }
    var table = [table_head];
    table.push(stemp)
    var ws = XLSX.utils.sheet_to_json(XLSX.utils.aoa_to_sheet(table));
    var html = `<table border="1" style="border-collapse:collapse"><tr><td>${Object.keys(ws[0]).join(`</td><td>`)}</td></tr><tr><td>${Object.values(ws[0]).join(`</td><td>`)}</td></tr></table>`;
    
    // 邮箱配置
    var mailOptions = {
      from: email, // sender address
      to: address, // 发送邮箱，以逗号隔开
      subject: 'node 小工具测试', // 主题
      text: 'node 小工具测试', // plain text body
      html, // html body,
    };
    var transporter = nodemailer.createTransport({
      host: 'smtp.chinatelecom.cn',
      port: 465,// SMTP 端口
      secure: true, // 使用 SSL
      auth: {
        user: email, // generated ethereal user
        pass: password 
      }
    });
    // loading[address] = true;
    document.getElementsByTagName('tr')[i].style.color = 'red';
    transporter.sendMail(mailOptions, (error, info, data) => {
      console.log(mailOptions.index);
      // let index = JSON.strigify(workbookJson).indexOf(info.accepted)
    // document.getElementsByTagName('tr')[mailOptions.index].style.color = 'green';
      if (error) {
        document.querySelector('#error').innerHTML = error.message
        return console.log(error);
      }
    // document.getElementsByTagName('tr')[i].style.color = 'green';
      console.log(info);
      console.log('Message sent: %s', info.messageId);
      transport.close();
    });
  }
}

