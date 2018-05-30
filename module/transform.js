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
    let values = Object.values(data);
    return { toEmail: values[values.length - 1], sheetJson: data }
  })
  return postJson;
}

module.exports = {
  transformToTable,
  transformToEmailJson
}