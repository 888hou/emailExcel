/*  require 一般全局 */
const nodemailer = require('nodemailer');
const XLSX = require('xlsx');
const { initMail, sendMail } = require('./module/mail');
const { transformToTable, transformToEmailJson } = require('./module/transform');


class PageHandler {
  constructor() {
    this.rABS = true; // true: readAsBinaryString ; false: readAsArrayBuffer
    this.workbookJson = [];
    this.rejected = [];
    this.file = null;
    this.buttonText = [
      { text: '确认发送邮件', disable: false, className: 'btn btn-default btn-lg' },
      { text: '正在发送', disable: true, className: 'btn btn-warning btn-lg' },
      { text: '重新发送失败列表', disable: false, className: 'btn btn-danger btn-lg' },
      { text: '发送成功', disable: true, className: 'btn btn-success btn-lg' }
    ];
    this.init();
  }

  init() {
    document.querySelector('#error').style.display = 'none';
    this.setButtonText(0);
    this.bindEvent();
  }

  bindEvent() {
    var self = this;

    document.querySelector('#xlf').addEventListener('change', self.handleFile.bind(self), false);

    document.querySelector('#submit').addEventListener('click', self.handleSubmit.bind(self), false);

    document.querySelector('#email').addEventListener('click', self.handleInput.bind(self), false);

    document.querySelector('#password').addEventListener('click', self.handleInput.bind(self), false);

  }

  setButtonText(status) {
    let { text, disable, className } = this.buttonText[status];
    let $submit = document.querySelector('#submit')
    $submit.innerHTML = text;
    $submit.disabled = disable;
    $submit.className = className;
  }

  handleFile(e) {
    var self = this;
    let { rABS, workbookJson, file } = self;
    self.init();
    // initial();
    var files = e.target.files;
    self.file = file = files[0];
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

  handleSubmit(e) {
    var self = this;
    let { workbookJson } = self;
    document.querySelector('#error').style.display = 'none';
    var email = document.querySelector('#email').value;
    var password = document.querySelector('#password').value;
    // 判断是否有被拒绝的列表
    var json = self.rejected.length ? self.rejected : workbookJson;
    var postJson = transformToEmailJson(json);
    this.send(email, password, postJson);
  }

  handleInput(e) {
    this.inputErrorOrSuccess(true)
  }

  inputErrorOrSuccess(success) {
    let inputs = document.querySelectorAll('.head-inputs input')
    for (var i = inputs.length - 1; i >= 0; i--) {
      inputs[i].style.borderColor = success ? '#dbd4e0' : '#d9534f';
    }
    document.querySelector('.input-error').style.display = success ? 'none' : 'block';
  }

  send(email, password, postJson) {
    var self = this;
    let { file } = self;
    self.setButtonText(1);
    console.log('发送', postJson);
    let fileName = file.name.split('.');
    fileName.pop();
    initMail(email, password).then(transporter => {
        let promises = postJson.map(data => {
          return sendMail(transporter, {
            ...data,
            subject: fileName.join('.'),
            fromEmail: email,
          })
        })
        return Promise.all(promises)
      }).then(data => {
        let rejectData = data.filter(item => item.code == 502);
        // 如果有发送错误的展示错误列表
        if (rejectData.length) {
          var rejects = rejectData.map(reject => {
            return { ...reject.data, '错误信息': reject.error }
          })
          self.rejected = rejectData.map(reject => reject.data);
          // 展示错误列表
          document.querySelector('#error').style.display = 'block';
          document.querySelector('#reject').innerHTML = transformToTable(rejects);
          self.setButtonText(2);
        } else {
          self.setButtonText(3);
        }
      })
      .catch(err => {
        self.setButtonText(0);
        if (err.code == 9) {
          self.inputErrorOrSuccess(false)
        }
        console.log(err)
      })
  }

}

window.onload = () => {
  new PageHandler();
}