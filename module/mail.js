// 发送邮件
const nodemailer = require('nodemailer');

// 初始化配置邮件
function initMail(email, password) {
  var promise = new Promise(function(resolve, reject) {
    if (!email || !password) {
      reject({ code: 9, msg: '初始化配置，邮件和密码不能为空' });
    } else {
      // 返回发送器
      var transporter = nodemailer.createTransport({
        host: 'smtp.chinatelecom.cn',
        port: 465, // SMTP 端口
        secure: true, // 使用 SSL
        auth: {
          user: email, // generated ethereal user
          pass: password
        }
      });
      resolve(transporter);
    }
  });

  return promise;
}

// 发送邮件
function sendMail(transporter, { fromEmail, toEmail, sheetJson }) {
  var html = `<table border="1" style="border-collapse:collapse"><tr><td>${Object.keys(sheetJson).join(`</td><td>`)}</td></tr><tr><td>${Object.values(sheetJson).join(`</td><td>`)}</td></tr></table>`;
  var mailOptions = {
    from: fromEmail, // sender address
    to: toEmail, // 发送邮箱，以逗号隔开
    subject: 'node 小工具测试', // 主题
    text: 'node 小工具测试', // plain text body
    html, // html body,
  };
  var promise = new Promise(function(resolve, reject) {
    transporter.sendMail(mailOptions, (error, info, data) => {

      if (error) {
        // 抛出异常
        resolve({ code: 502, msg: '邮箱发送失败', data: sheetJson, error })
      } else {
        transporter.close();
        resolve({ code: 200, msg: '邮箱发送成功', data: sheetJson, info })
      }
    });
  })

  return promise;

}

module.exports = {
  initMail,
  sendMail
};