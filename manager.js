window.onload = () => {


const nodemailer = require('nodemailer');
const XLSX = require('xlsx')
var xlf = document.querySelector('#xlf');
var rABS = true; // true: readAsBinaryString ; false: readAsArrayBuffer
var workbook;
var workbookJson;
var number;
var accepted = [];
var rejected;

function initial() {
  number = 0;
  // accepted = [];
  rejected = [];
  document.querySelector('#reject').innerHTML = '';
  document.querySelector('#error').innerHTML = '';
}

function handleFile(e) {

  var files = e.target.files,
    f = files[0];
  document.querySelector('.icon').className='glyphicon glyphicon-folder-close icon'
  document.querySelector('.head-icon-text').innerHTML = f.name;
  var reader = new FileReader();
  reader.onload = function(e) {
    var data = e.target.result;
    if (!rABS) data = new Uint8Array(data);
    // 读取表格
    workbook = XLSX.read(data, {
      type: rABS ? 'binary' : 'array'
    });
    // 保存表格数组
    workbookJson = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
    // 页面打印表格
    var html1 = ''
    workbookJson.map(data => {
      html1 += `<tr><td>${Object.values(data).join('</td><td>')}</td></tr>`;
    })
    var html = `<table border="1" style="border-collapse:collapse">
        <tr><td>${Object.keys(workbookJson[0]).join(`</td><td>`)}</td></tr>
        ${html1}
      </table>`;
    document.querySelector('#excel').innerHTML = html;
  };
  if (rABS) {
    reader.readAsBinaryString(f)
  } else {
    reader.readAsArrayBuffer(f)
  };
}
xlf.addEventListener('change', handleFile, false);


submit = function(e) {
  document.querySelector('#submit').innerHTML = '正在处理';
  initial();
  var email = document.querySelector('#email').value;
  var password = document.querySelector('#password').value;
  transfer(workbook, email, password)
}

// 将表格每行信息整理出来
function transfer(workbook, email, password) {
  // 获取第一个sheet名
  var first_sheet_name = workbook.SheetNames[0];
  var address_of_cell = 'A1';
  // 获取第一个sheet
  var worksheet = workbook.Sheets[first_sheet_name];
  var range = XLSX.utils.decode_range(worksheet['!ref']);
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

    send(email, password, address, ws, workbookJson.length);
    document.getElementsByTagName('tr')[i].style.color = 'yellow';

  }
}

// 发送邮件
function send(email, password, receiveAddress, sheetJson, postNum) {
  console.log(postNum)
  var html = `<table border="1" style="border-collapse:collapse"><tr><td>${Object.keys(sheetJson[0]).join(`</td><td>`)}</td></tr><tr><td>${Object.values(sheetJson[0]).join(`</td><td>`)}</td></tr></table>`;
  // 邮箱配置
  var mailOptions = {
    from: email, // sender address
    to: receiveAddress, // 发送邮箱，以逗号隔开
    subject: 'node 小工具测试', // 主题
    text: 'node 小工具测试', // plain text body
    html, // html body,
  };
  var transporter = nodemailer.createTransport({
    host: 'smtp.chinatelecom.cn',
    port: 465, // SMTP 端口
    secure: true, // 使用 SSL
    auth: {
      user: email, // generated ethereal user
      pass: password
    }
  });

  console.log('发送的邮箱是==>', receiveAddress);
  console.log('邮箱密码==>', email, password);
  document.querySelector('#submit').innerHTML = '正在发送';

  transporter.sendMail(mailOptions, (error, info, data) => {

    number++;

    if (error) {
      // TODO 报错之后获取不到是哪一条的报错
      console.log(info);
      let div = document.createElement("div");
      div.innerHTML = error.message;
      document.querySelector('#error').append(div)
      console.log('错误')
    }
    // 将发送成功的存入accepted数组
    info && accepted.push(info.accepted[0]);
    // 如果全部发送完毕
    if (number == postNum) {
      checkRejected()
    }
    if (info) {
      console.log('Message sent: %s', info.messageId);
    }
    transporter.close();
  });
}

// 查看哪些是报错的
function checkRejected() {
  workbookJson.map((item, index) => {
    let value = Object.values(item);
    let address = value[value.length - 1];
    if (accepted.indexOf(address) == -1) {
      rejected.push(index)
      document.getElementsByTagName('tr')[index + 1].style.color = 'red';
    } else {
      document.getElementsByTagName('tr')[index + 1].style.color = 'green';
    }
  })
  if (rejected.length) {
    let html2 = '';
    rejected.map(i => {
      html2 += `<tr><td>${Object.values(workbookJson[i]).join(`</td><td>`)}</td></tr>`
    })
    document.querySelector('#reject').innerHTML = `<div class="excel"><div class="content-title">发送失败列表</div><table border="1" style="border-collapse:collapse">
        <tr><td>${Object.keys(workbookJson[0]).join(`</td><td>`)}</td></tr>
        ${html2}
      </table></div>
      <button onclick="resubmit()" class="btn btn-warning btn-lg">重新发送失败列表</button>`
    document.querySelector('#submit').innerHTML = '确认发送邮件';
  } else {
    document.querySelector('#submit').innerHTML = '发送成功';
  }
}

resubmit = function() {
  let rejectedList = rejected;
  initial();
  var email = document.querySelector('#email').value;
  var password = document.querySelector('#password').value;
  rejectedList.map(index => {

    let data = workbookJson[index];
    let keys = Object.keys(data)
    let address = data[keys[keys.length - 1]];
    send(email, password, address, [data], rejectedList.length);

  })
}

}