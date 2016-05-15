var convert = require('./lib/convert.js');
var co = require('co');
var thunkify = require('thunkify');
var process = thunkify(convert.process);

// convert.info('./upload/dev.xlsx', function(err, path) {
//   if (err) {
//     return console.log(err);
//   }
//   console.log(path);
// })

co(function*() {
  var test1 = yield process('./upload/dev.xlsx', ['款号', '品牌'], '季节', '季节组');
  console.log(test1);
})
return


convert.process('./upload/dev.xlsx', ['款号', '品牌'], '季节', '季节组', function(err, info) {
  if (err) {
    return console.log(err);
  }
  console.log(info);
})
