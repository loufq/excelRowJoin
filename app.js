/**
 * Module dependencies.
 */

var render = require('./lib/render');
var convert = require('./lib/convert.js');
var logger = require('koa-logger');
var route = require('koa-route');
var serve = require('koa-static');
var fs = require('fs-extra')
var os = require('os');
var path = require('path');
var parse = require('co-body');
var thunkify = require('thunkify');
var parsebusboy = require('co-busboy');
var urlparamparser = require('url-param-parser');
var koa = require('koa');
var app = koa();
var body = require('koa-better-body')
  // "database"
var posts = [];
// middleware
app.use(body())
app.use(logger());
//使用静态文件
app.use(serve('./scripts'));
app.use(serve('./resource'));
app.use(serve('./upload'));

// route middleware
app.use(route.get('/', list));
app.use(route.post('/fileupload', fileupload));
app.use(route.get('/process', process));

function* fileupload() {
  var files = this.body.files;
  var file = files.file;
  var uploadPath = file.path;
  console.log(file);
  var descPath = './upload/' + file.name
  fs.copySync(uploadPath, descPath)
  var xlsInfo = convert.info(descPath)
  console.log(xlsInfo);
  //处理转化成excel-输出
  this.body = xlsInfo;
  //this.redirect('./1.xlsx');
}


function* process() {
  var info = urlparamparser(decodeURI(this.req.url)).search
  var path = info.path.replace(/(\s*)|(\s*)/g, "")
  var sheetname = info.sheetname.replace(/(\s*)|(\s*)/g, "")
  var compareNames2 = info.compareNames.replace(/(\s*)|(\s*)/g, "").split(',')
  var joincoldistinct = true
  if (info.joincoldistinct == 'true') {
    joincoldistinct = true
  } else {
    joincoldistinct = false
  }
  var compareNames = []
  for (var i = 0; i < compareNames2.length; i++) {
    var str = compareNames2[i]
    if (str.length) {
      compareNames.push(str)
    }
  }
  var joinName = info.joinName.replace(/(\s*)|(\s*)/g, "")
  var joinedName = info.joinedName.replace(/(\s*)|(\s*)/g, "")
  var cfun = thunkify(convert.process);
  var results = yield cfun(path, compareNames, joinName, joinedName, joincoldistinct);
  this.redirect(results.outputurl);
}

/**
 * Post listing.
 */

function* list() {
  this.body = yield render('list', {
    posts: posts
  });
}

// listen

app.listen(3000);
console.log('listening on port 3000');
