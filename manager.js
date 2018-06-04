window.onload = () => {
  const nodemailer = require('nodemailer');
  const XLSX = require('xlsx');
  const { initMail, sendMail } = require('./module/mail');
  const { transformToTable, transformToEmailJson } = require('./module/transform');

  var rABS = true; // true: readAsBinaryString ; false: readAsArrayBuffer
  var workbookJson;
  var rejected = [];
  var file;
  var buttonText = [
    { text: '确认发送邮件', disable: false, className: 'btn btn-default btn-lg' },
    { text: '正在发送', disable: true, className: 'btn btn-warning btn-lg' },
    { text: '重新发送失败列表', disable: false, className: 'btn btn-danger btn-lg' },
    { text: '发送成功', disable: true, className: 'btn btn-success btn-lg' }
  ];

  function setButtonText(status) {
    let { text, disable, className } = buttonText[status];
    document.querySelector('#submit').innerHTML = text;
    document.querySelector('#submit').disabled = disable;
    document.querySelector('#submit').className = className;
  }

  // 重置
  function initial() {
    rejected = [];
    document.querySelector('#error').style.display = 'none';
    setButtonText(0);
  }

  document.querySelector('#xlf').addEventListener('change', handleFile, false);

  function handleFile(e) {
    initial();
    var files = e.target.files;
    file = files[0];
    document.querySelector('.icon').className = 'glyphicon glyphicon-folder-close icon'
    document.querySelector('.head-icon-text').innerHTML = file.name;
    var reader = new FileReader();
    reader.onload = function(e) {
      var data = e.target.result;
      if (!rABS) data = new Uint8Array(data);
      // 读取表格
      let workbook = XLSX.read(data, {
        type: rABS ? 'binary' : 'array'
      });
      // 保存表格数组
      workbookJson = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
      // 页面打印表格
      document.querySelector('#excel').innerHTML = transformToTable(workbookJson);
    };
    if (rABS) {
      reader.readAsBinaryString(file)
    } else {
      reader.readAsArrayBuffer(file)
    };
  }

  // 是否显示邮箱密码错误提示
  function inputErrorOrSuccess(success) {
    let inputs = document.querySelectorAll('.head-inputs input')
    for (var i = inputs.length - 1; i >= 0; i--) {
      inputs[i].style.borderColor = success ? '#dbd4e0' : '#d9534f';
    }
    document.querySelector('.input-error').style.display = success ? 'none' : 'block';
  }

  changeInput = function(e) {
    inputErrorOrSuccess(true)
  }

  submit = function(e) {
    document.querySelector('#error').style.display = 'none';
    var email = document.querySelector('#email').value;
    var password = document.querySelector('#password').value;
    // 判断是否有被拒绝的列表
    var json = rejected.length ? rejected : workbookJson;
    var postJson = transformToEmailJson(json);
    send(email, password, postJson);
  }

  // 发送邮件
  function send(email, password, postJson) {
    setButtonText(1);
    console.log('发送', postJson);
    let fileName = file.name.split('.');
    fileName.pop();
    initMail(email, password).then(transporter => {
        let promises = postJson.map(data => {
          return sendMail(transporter, {
            ...data,
            subject:fileName.join('.'), 
            fromEmail: email,
          })
        })
        return Promise.all(promises)
      }).then(data => {
        let rejectData = data.filter(item => item.code == 502);
        // 如果有发送错误的展示错误列表
        if (rejectData.length) {
          rejects = rejectData.map(reject => {
            return { ...reject.data, '错误信息': reject.error }
          })
          rejected = rejectData.map(reject => reject.data);
          // 展示错误列表
          document.querySelector('#error').style.display = 'block';
          document.querySelector('#reject').innerHTML = transformToTable(rejects);
          setButtonText(2);
        } else {
          setButtonText(3);
        }
      })
      .catch(err => {
        setButtonText(0);
        if (err.code == 9) {
          inputErrorOrSuccess(false)
        }
        console.log(err)
      })
  }

}