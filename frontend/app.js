var bucketName = 'BUCKET_NAME';
var bucketRegion = 'REGION';
var IdentityPoolId = 'IDENTITY_POOL_ID';

AWS.config.update({
  region: bucketRegion,
  credentials: new AWS.CognitoIdentityCredentials({
    IdentityPoolId: IdentityPoolId
  })
});

var s3 = new AWS.S3({
  apiVersion: '2006-03-01',
  params: {Bucket: bucketName}
});

function formatBytes(bytes,decimals) {
   if(bytes === 0) return '0 Bytes';
   var k = 1000,
       dm = decimals || 2,
       sizes = ['Bytes', 'KB', 'MB', 'GB', 'TB', 'PB', 'EB', 'ZB', 'YB'],
       i = Math.floor(Math.log(bytes) / Math.log(k));
   return parseFloat((bytes / Math.pow(k, i)).toFixed(dm)) + ' ' + sizes[i];
}

function listINVs() {
  s3.listObjects({Delimiter: '/',Prefix: 'output/',Marker: 'output/',}, function(err, data) {
    if (err) {
      return alert('There was an error listing your invoices: ' + err.message);
    }
    var href = this.request.httpRequest.endpoint.href;
    var bucketUrl = href + bucketName + '/output/';

    var files = data.Contents.map(function(file) {
      var fileKey = file.Key.replace(data.Prefix, '');
      var fileUrl = bucketUrl + encodeURIComponent(fileKey);
      var fileSize = formatBytes(file.Size);
      var fileDate = file.LastModified;
      var options = {month: 'numeric',day: 'numeric',year: '2-digit',hour: 'numeric',minute: 'numeric',hour12: true,};
      var fileTime = fileDate.toLocaleString('en-US', options);
      return getHtml([
        '<tr>',
          '<td><a href="' + fileUrl + '">' + fileKey + '</a></td>',
          '<td>' + fileSize + '</td>',
          '<td>' + fileTime + '</td>',
        '</tr>',
      ]);
    });
    var message = files.length ?
      '<p>Click a file name to download an invoice file.</p>':
      '<p>There are no invoice files.</p>';
    var htmlTemplate = ([
      '<h2>Converted Invoice Files</h2>',
      message,
      '<table>',
        '<tr>',
          '<th>File Name</th>',
          '<th>File Size</th>',
          '<th>Last Modified</th>',
        '</tr>',
        getHtml(files),
      '</table>',
      '<h2>Upload a Spreadsheet</h2>',
      '<p>Uploaded spreadsheets must:</p>',
      '<ul>',
        '<li>follow strict template formatting.</li>',
        '<li>be in .xlsx format.</li>',
        '<li>contain 500 or fewer line items.</li>',
      '</ul>',
        '<input id="fileupload" type="file" accept=".xlsx">',
        '<button id="addfile" onclick="uploadXLSX()">',
          'Upload Spreadsheet',
        '</button>',
      '<p> If converted file doesn\'t appear automatically, <a href="javascript:location.reload(true)">reload page</a>.</p>',
      '<p class="disclaimer"><em>If converted file doesn\'t appear after reload, there may be a problem with your spreadsheet.</em></p>',
    ]);
    document.getElementById('app').innerHTML = getHtml(htmlTemplate);
  });
}

function uploadXLSX() {
  var files = document.getElementById('fileupload').files;
  if (!files.length) {
    return alert('Please choose a file to upload first.');
  }
  var file = files[0];
  var sheetKey = 'input/' + file.name;

  s3.upload({
    Key: sheetKey,
    Body: file,
    ACL: 'public-read'
  }, function(err, data) {
    if (err) {
      return alert('There was an error uploading your spreadsheet: ', err.message);
    }
    alert('Successfully uploaded spreadsheet.');
    listINVs();
  });
}
