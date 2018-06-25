function transformToTable(sheetJson) {
  var html1 = ''
  sheetJson.map(data => {
    html1 += `<tr><td>${Object.values(data).join('</td><td>')}</td></tr>`;
  })
  var html = `<table border="1" style="border-collapse:collapse">
        <tr><td>${Object.keys(sheetJson[0]).join(`</td><td>`)}</td></tr>
        ${html1}
      </table>`;
  return html;
}

function transformToEmailJson(workbookJson) {
  var postJson = workbookJson.map(data => {
    // let values = Object.values(data);
    // var json = data;
    let keys = Object.keys(data);
    let emailKey = keys[keys.length -1];
    let toEmail = data[emailKey];
    delete data[emailKey];
    return { toEmail, sheetJson: data, originJson: {...data, [emailKey]: toEmail} }
  })
  return postJson;
}

module.exports = {
  transformToTable,
  transformToEmailJson
}