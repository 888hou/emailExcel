var {initMail, sendMail} = require('../module/mail');

initMail('liwenbin@chinatelecom.cn', 'liwenbin').then((transporter) => {
    var dataList = [{
        fromEmail: '837531387@qq.com', toEmail: '666@126.com', sheetJson: 'hehe'
    }, {
        fromEmail: '837531387@qq.com', toEmail: '666@126.com', sheetJson: 'hehe'
    }, {
        fromEmail: '837531387@qq.com', toEmail: '666@126.com', sheetJson: 'hehe'
    }, {
        fromEmail: '837531387@qq.com', toEmail: '666@126.com', sheetJson: 'hehe'
    }];
    var promiseList = dataList.map((item) => {
        return sendMail(transporter, item)
    });

    return Promise.all(promiseList)
}).then((data) => {
    console.log(data)
}).catch((err) => {
    console.log(err);
});