/**
 * Module dependencies.
 */

xlsx2json = require('xlsx-json')
var xlsx = require('xlsx')
  //重新导出
var xls = require('node-simple-xlsx');

Array.prototype.contains  =   function(item) {    
  return  RegExp(item).test(this);
};
Array.prototype.unique = function() {
  this.sort();
  var re = [this[0]];
  for (var i = 1; i < this.length; i++) {
    if (this[i] !== re[re.length - 1]) {
      re.push(this[i]);
    }
  }
  return re;
}


function clone(myObj) {
  if (typeof(myObj) != 'object') return myObj;
  if (myObj == null) return myObj;
  var myNewObj = new Object();
  for (var i in myObj)
    myNewObj[i] = clone(myObj[i]);
  return myNewObj;
}

var info2 = function(filePath) {
  var workbook = xlsx.readFile(filePath)
  var sheetNames = workbook.SheetNames
  var first_sheet_name = sheetNames[0]
  var worksheet = workbook.Sheets[first_sheet_name]
  var address_of_cellIndex = 0
  var address_of_cellColT = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T']
  var needContiune = true
  var colTitles = []
  while (needContiune) {
    var address_of_cell = address_of_cellColT[address_of_cellIndex] + '1';
    var desired_cell = worksheet[address_of_cell];
    if (!desired_cell) {
      break;
    }
    var desired_value = desired_cell.v;
    if (desired_value.length == 0) {
      needContiune = false;
    } else {
      colTitles.push({
        'cell': address_of_cell,
        'cellColIndex': address_of_cellColT[address_of_cellIndex],
        'name': desired_value
      })
    }
    address_of_cellIndex++
  }
  var info = {
    'path': filePath,
    'sheetname': first_sheet_name,
    'titles': colTitles
  }
  return info
}

var process2 = function(filePath, compareNames, joinName, joinedName, joincoldistinct, cb) {

  var info = info2(filePath)
  var workbook = xlsx.readFile(filePath)
  var sheet_name = info.sheetname
  var worksheet = workbook.Sheets[sheet_name]
  var line = 2;
  //读取一行数据
  var needContiune = true
  var allRowData = []
  while (needContiune) {
    var lineData = []
    var lineNeedJoin = true
    info.titles.forEach(function(item, index) {
      var address_of_cell = item['cellColIndex'] + '' + line
      var desired_cell = worksheet[address_of_cell];
      if (index == 0 && (!desired_cell || (desired_cell.v).length == 0)) {
        lineNeedJoin = false
        needContiune = false;
        return
      }
      if (lineNeedJoin) {
        var val = ''
        if (!desired_cell) {
          val = ''
        } else {
          var desired_value = desired_cell.v;
          val = desired_value
        }
        lineData.push({
          'col': index,
          'row': (line - 1),
          'cell': address_of_cell,
          'title': item['name'],
          'val': val
        })
      }

    })
    if (lineData.length) {
      allRowData.push(lineData)
    }
    line++
  }
  var afterAllData = []
  allRowData.forEach(function(item, index) {
    //找相似的
    var compareBase = ''
    item.forEach(function(item2, index) {
      if (compareNames.contains(item2.title)) {
        compareBase += item2.val
      }
    })

    //结果里找当前行有没有已经存在的
    var hasOri = false
    afterAllData.forEach(function(item3, index) {
      var compareFor = ''
      item3.forEach(function(item4, index) {
        if (compareNames.contains(item4.title)) {
          compareFor += item4.val
        }
      })

      //已经存在
      if (compareFor == compareBase) {
        //找到joinedNameItem
        var joinedNameItem = null
        item3.forEach(function(item2, index) {
          if (item2['title'] != null && item2['title'] == joinedName) {
            joinedNameItem = item2
          }
        })

        var joinNameItem = null
        item.forEach(function(item2, index) {
          if (item2['title'] != null && item2['title'] == joinName) {
            joinNameItem = item2
          }
        })
        hasOri = true
        joinedNameItem.val.push(joinNameItem['val'])
        return false;
      }
    })

    //compareNames, joinName, joinedName
    if (!hasOri) { //不存在，创建
      //找到joinName的值
      var joinNameItem = null
      item.forEach(function(item2, index) {
        if (item2['title'] != null && item2['title'] == joinName) {
          joinNameItem = item2
        }
      })
      var newItem = []
      item.forEach(function(item2, index) {
        newItem.push(item2)
      })
      newItem.push({
        'col': 0,
        'row': joinNameItem['row'],
        'cell': '',
        'title': joinedName,
        'val': [joinNameItem['val']]
      })
      afterAllData.push(newItem)
    }

  })

  afterAllData.forEach(function(item, index) {
    var joinedNameItem = null
    item.forEach(function(item2, index) {
      if (item2['title'] != null && item2['title'] == joinedName) {
        joinedNameItem = item2
      }
    })
    if (joincoldistinct) {
      joinedNameItem['val'] = joinedNameItem['val'].unique().join(',')
    } else {
      joinedNameItem['val'] = joinedNameItem['val'].join(',')
    }
  })

  var output = []
  afterAllData.forEach(function(itemOri, index) {
    //数组转键值对
    var outputLine = {}
    itemOri.forEach(function(cellItem, index) {
      outputLine[cellItem['title']] = cellItem['val']
    })
    output.push(outputLine)
  })

  var outputpath = './upload/' + new Date().getTime() + '.xlsx'
    //var outputpath = './' + new Date().getTime() + '.xlsx'
  var outputurl = './' + new Date().getTime() + '.xlsx'
  info['outpath'] = outputpath
  info['outputurl'] = outputurl
  xls.write(outputpath, output, function(err) {
    if (err) {
      cb(err)
    } else {
      cb(null, info)
    }
  });

}


module.exports.info = info2;

module.exports.process = process2
