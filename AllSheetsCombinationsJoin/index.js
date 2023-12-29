/*
 * @Author: Wade Zhong wzhong@hso.com
 * @Date: 2023-12-29 15:17:35
 * @LastEditTime: 2023-12-29 15:50:36
 * @LastEditors: Wade Zhong wzhong@hso.com
 * @Description: 获取当前工作簿的所有sheet。把所有sheet的数据都笛卡尔积出数据，输出到指定sheetName里面（有就清空添加、没有就添加）
 * @FilePath: \JavaScriptWPSMacros\AllSheetsCombinationsJoin\index.js
 * Copyright (c) 2023 by Wade Zhong wzhong@hso.com, All Rights Reserved.
 */
//let Workbook = ThisWorkbook //获取当前代码所在的工作簿对象
//let sht1 = wb.ActiveSheet //获取当前显示的工作表（活动工作表）
//let sht2 = wb.Sheets('Sheet1') //获取当前名为Sheet1的工作表


function cartesianProduct(arr) {
    return arr.reduce(function(a, b) {
        return a.map(function(x) {
            return b.map(function(y) {
                return x.concat(y);
            })
        }).reduce(function(a, b) { return a.concat(b) }, [])
    }, [[]])
}

function joinAllSheets() {
    var resultSheetName = "Result"; // 保存结果表的名称，方便修改

    var sheetNames = [];
    for (var i = 1; i <= ThisWorkbook.Sheets.Count; i++) {
        var sheetName = ThisWorkbook.Sheets(i).Name;
        if (sheetName !== resultSheetName) { // 排除结果表
            sheetNames.push(sheetName);
        }
    }

    var sheetData = sheetNames.map(function(name) {
        var sheet = ThisWorkbook.Sheets(name);
        var lastColumn = sheet.Cells(1, sheet.Columns.Count).End(-4159).Column; // 获取最后一列的列号
        var lastRow = sheet.Cells(sheet.Rows.Count, 1).End(-4162).Row; // 获取最后一行的行号
        var lastColumnLetter = String.fromCharCode(64 + lastColumn); // 将列号转换为字母
        return sheet.Range("A1:" + lastColumnLetter + lastRow).Value2; // 获取表中所有的数据
    });

    var result = cartesianProduct(sheetData);

    var resultSheet;
    try {
        resultSheet = ThisWorkbook.Sheets(resultSheetName); // 尝试获取结果表
        resultSheet.Cells.Clear(); // 清空结果表
    } catch (e) {
        resultSheet = ThisWorkbook.Sheets.Add(); // 如果结果表不存在，则创建一个新的表
        resultSheet.Name = resultSheetName; // 将新表命名为结果表的名称
    }

    for (var i = 0; i < result.length; i++) {
        for (var j = 0; j < result[i].length; j++) {
            resultSheet.Cells(i + 1, j + 1).Value2 = result[i][j];
        }
    }
}
