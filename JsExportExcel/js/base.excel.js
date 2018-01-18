/**
 * 
 * Name base.excel
 * Description 浏览器端js下载table,json到excel示例
 * 1.先判断浏览器类型:IE和非IE浏览器
 * IE流量器:使用ie命令方式,将html或是csv输出到open的window，然后使用execCommand的saveas命令，存为csv或xls。
 * 非IE浏览器(支持data协议）：将html或是csv先用js encodeURIComponent处理，然后前缀data:text/xls;charset=utf-8,\ufeff，即可使浏览器将其中的数据当做excel来处理，浏览器将提示下载或打开excel文件。
 * 2.后缀为xls的文件保持原表格样式（支持跨行、跨列），后缀为.csv的文件不支持支持跨行、跨列，但支持更多程序打开，
 * 请根据实际情况选择导出文件类型(.xls或csv)
 * Author heyp
 * create Date:2018-01-15
 * 
 */


/**
*判断是否IE浏览器
*是--true,否--false
**/
function isIE() { //ie?
    if (!!window.ActiveXObject || "ActiveXObject" in window) {
        return true;
    }
    else {
        return false;
    }
};

/*
*获取表格内容,IE和其他浏览器分开处理
*
*/
var getTblData = {
    IE: function (tableId, inWindow) {
        var rows = 0,
            tblDocument = document;
        if (!!inWindow && inWindow != "") {
            if (!document.all(inWindow)) {
                return null;
            } else {
                tblDocument = eval(inWindow).document;
            }
        }
        var curTbl = tblDocument.getElementById(tableId);
        if (curTbl.rows.length > 65000) {
            alert('源行数不能大于65000行');
            return false;
        }
        if (curTbl.rows.length <= 1) {
            alert('数据源没有数据');
            return false;
        }
        var outStr = "";
        if (curTbl != null) {
            var itemCell = null,
                self = this,
                exportTableObj = self.rowSpanCellSplit(curTbl);
            for (var j = 0; j < exportTableObj.rows.length; j++) {
                for (var i = 0; i < exportTableObj.rows[j].cells.length; i++) {
                    itemCell = exportTableObj.rows[j].cells[i];
                    if (j > 0 && itemCell.hasChildNodes() && itemCell.firstChild.nodeName.toLowerCase() == "input") {
                        if (itemCell.firstChild.type.toLowerCase() == "checkbox") {
                            if (itemCell.firstChild.checked == true) {
                                outStr += "是" + "\t";
                            } else {
                                outStr += "否" + "\t";
                            }
                        }
                    } else {
                        outStr += " " + itemCell.innerText.replace(/,/g, '') + "\t";
                    }
                    if (itemCell.colSpan > 1) {
                        for (var k = 0; k < itemCell.colSpan - 1; k++) {
                            outStr += " \t";
                        }
                    }
                }
                outStr += "\r\n";
            }
        } else {
            outStr = null;
            alert(tableId + "不存在!");
        }
        return outStr;
    },
    other: function (tableId, inWindow) {
        var rows = 0,
            tblDocument = document;
        var curTbl = tblDocument.getElementById(tableId);
        var outStr = "";
        if (curTbl != null) {
            var itemCell = null,
                self = this,
                exportTableObj = self.rowSpanCellSplit(curTbl);
            for (var j = 0; j < exportTableObj.rows.length; j++) {
                for (var i = 0; i < exportTableObj.rows[j].cells.length; i++) {
                    itemCell = exportTableObj.rows[j].cells[i];
                    outStr += itemCell.innerText.replace(/,/g, '') + ","; //去掉逗号";
                    if (itemCell.colSpan > 1) {
                        for (var k = 0; k < itemCell.colSpan - 1; k++) {
                            outStr += ","; // \t";
                        }
                    }
                }
                outStr += "\r\n";
            }
        }
        else {
            outStr = null;
            alert(tableId + "不存在 !");
        }
        return outStr;
    },
    rowSpanCellSplit: function (curTbl) {
        var exportTableObj = $('<table>' + $(curTbl).html() + '</table>')[0];
        var maxColumns = 0;
        $.each(exportTableObj.rows[0].cells, function (i, item) {
            maxColumns++;
            if (item.rowSpan > 1) {
                maxColumns += item.rowSpan - 1;
            }
        });
        var itemCell = null,
            columnIndex = 0;
        //将跨行单元格拆分
        for (var i = 0; i < maxColumns; i++) {
            for (var j = 0; j < exportTableObj.rows.length; j++) {
                itemCell = exportTableObj.rows[j].cells[i];
                if (itemCell && itemCell.rowSpan > 1) {
                    columnIndex = 0;
                    $(itemCell).prevAll().each(function (k, item) {
                        columnIndex++;
                        if (item.colSpan > 1) {
                            columnIndex += item.colSpan - 1;
                        }
                    });
                    for (var k = itemCell.rowSpan - 1; k > 0; k--) {
                        if (i == exportTableObj.rows[j].cells.length - 1) {
                            $(exportTableObj.rows[j + k].cells[columnIndex - 1]).after('<' + itemCell.tagName + ' colSpan="' + itemCell.colSpan + '"></th>');
                        } else {
                            $(exportTableObj.rows[j + k].cells[columnIndex]).before('<' + itemCell.tagName + ' colSpan="' + itemCell.colSpan + '"></th>');
                        }
                    }
                    $(itemCell).attr('rowSpan', 1);
                }
                if (itemCell) {
                    itemCell.innerText = itemCell.innerText.replace(/[\r\n]/g, "")//去掉回车换行
                }
            }
        }
        return exportTableObj;
    }
};

/*
*转换json数组为excelStr
*/
var exchangeJsonToExcelStr = {
    XLS: function (json) {
        var outStr = '',
            theadHtml = '<thead><tr>';
        //表头
        for (var key in json[0]) {
            theadHtml += '<th>' + key + '</th>';
        }
        theadHtml += '</tr></thead>';
        //内容
        var tbodyHtml = '<tbody>';
        $.each(json, function (i, item) {
            tbodyHtml += '<tr>';
            for (var key in item) {
                tbodyHtml += '<td>' + item[key] + '</td>';
            }
            tbodyHtml += '</tr>';
        });
        tbodyHtml += '</tbody>';
        outStr = '<table>' + theadHtml + tbodyHtml + '</table>';
        return outStr;
    },
    CSV: function (json) {
        var _self = exchangeJsonToExcelStr;
        if (isIE()) {
            return _self.CSV_IE(json);
        } else {
            return _self.CSV_Other(json);
        }
    },
    CSV_IE: function (json) {
        var outStr = '';
        //表头
        for (var key in json[0]) {
            outStr += ' ' + key + '\t';
        }
        outStr += '\r\n';
        //内容
        $.each(json, function (i, item) {
            for (var key in item) {
                outStr += ' ' + item[key] + '\t';
            }
            outStr += '\r\n';
        });
        return outStr;
    },
    CSV_Other: function (json) {
        var outStr = '';
        //表头
        for (var key in json[0]) {
            outStr += key + ',';
        }
        outStr += '\r\n';
        //内容
        $.each(json, function (i, item) {
            for (var key in item) {
                outStr += item[key] + ',';
            }
            outStr += '\r\n';
        });
        return outStr;
    }
};

/**
 * 导出excel为xls
 */
var exportExcelToXls = {
    exportFromTable: function (tableId, fileName, inWindow) {
        var allStr = '',
            curStr = '';
        curStr = $('#' + tableId).html();
        if (curStr != null && curStr != '') {
            allStr = '<table>' + curStr + '</table>';
        }
        else {
            alert("你要导出的表不存在！");
            return false;
        }
        exportExcel.doExport(fileName, allStr, '.xls');
    },
    expoprtFromJson: function (json, fileName, inWindow) {
        var allStr = '',
            curStr = '';
        if (!json || !json.length) {
            alert('数据源没有数据');
            return false;
        }
        curStr = exchangeJsonToExcelStr.XLS(json);
        if (curStr != null) {
            allStr += curStr;
        }
        else {
            alert("数据源没有数据！");
            return false;
        }
        exportExcel.doExport(fileName, allStr, '.xls');
    }
};

/**
 * 导出excel为csv
 */
var exportExcelToCSV = {
    exportFromTable: function (tableId, fileName, inWindow) {
        var allStr = '',
            curStr = '';
        if (isIE()) { //如果是IE浏览器  
            curStr = getTblData.IE(tableId, inWindow);
        } else {
            curStr = getTblData.other(tableId, inWindow);
        }
        if (curStr != null) {
            allStr += curStr;
        }
        else {
            alert("你要导出的表不存在！");
            return false;
        }
        exportExcel.doExport(fileName, allStr, '.csv');
    },
    expoprtFromJson: function (json, fileName, inWindow) {
        var allStr = '',
            curStr = '';
        if (isIE()) { //如果是IE浏览器  
            curStr = exchangeJsonToExcelStr.CSV_IE(json);
        } else {
            curStr = exchangeJsonToExcelStr.CSV_Other(json);
        }
        if (curStr != null) {
            allStr += curStr;
        }
        else {
            alert("数据源没有数据！");
            return false;
        }
        exportExcel.doExport(fileName, allStr, '.csv');
    }
};

/**
 * 执行excel导出
 */
var exportExcel = {
    exportFromTable: function (tableId, fileName, inWindow) {
        if (!tableId || tableId == "null") {
            alert("你要导出的表不存在！");
            return false;
        }
        var _self = exportExcel;
        fileName = _self.getFileNameAndExtension(fileName);
        var extension = fileName.substr(fileName.lastIndexOf('.'));
        if (extension == '.csv') {
            exportExcelToCSV.exportFromTable(tableId, fileName, inWindow);
        } else {
            exportExcelToXls.exportFromTable(tableId, fileName, inWindow);
        }
    },
    expoprtFromJson: function (json, fileName, inWindow) {
        if (!json || !json.length) {
            alert('数据源没有数据');
            return false;
        }
        var _self = exportExcel;
        fileName = _self.getFileNameAndExtension(fileName);
        var extension = fileName.substr(fileName.lastIndexOf('.'));
        if (extension == '.csv') {
            exportExcelToCSV.expoprtFromJson(json, fileName, inWindow);
        } else {
            exportExcelToXls.expoprtFromJson(json, fileName, inWindow);
        }
    },
    doExport: function (fileName, excelStr, extension) {
        var _self = exportExcel;
        if (excelStr == null) {
            alert("导出数据源不存在！");
            return false;
        }
        fileName = _self.getFileNameAndExtension(fileName);
        var extension = fileName.substr(fileName.lastIndexOf('.'));
        var excelFileStr = '';
        if (extension == '.xls') {
            excelFileStr = "<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:x='urn:schemas-microsoft-com:office:excel' xmlns='http://www.w3.org/TR/REC-html40'>";
            excelFileStr += '<meta http-equiv="content-type" content="application/vnd.ms-excel; charset=UTF-8">';
            excelFileStr += '<meta http-equiv="content-type" content="application/vnd.ms-excel';
            excelFileStr += '; charset=UTF-8">';
            excelFileStr += "<head>";
            excelFileStr += "<!--[if gte mso 9]>";
            excelFileStr += "<xml>";
            excelFileStr += "<x:ExcelWorkbook>";
            excelFileStr += "<x:ExcelWorksheets>";
            excelFileStr += "<x:ExcelWorksheet>";
            excelFileStr += "<x:Name>";
            excelFileStr += "{worksheet}";
            excelFileStr += "</x:Name>";
            excelFileStr += "<x:WorksheetOptions>";
            excelFileStr += "<x:DisplayGridlines/>";
            excelFileStr += "</x:WorksheetOptions>";
            excelFileStr += "</x:ExcelWorksheet>";
            excelFileStr += "</x:ExcelWorksheets>";
            excelFileStr += "</x:ExcelWorkbook>";
            excelFileStr += "</xml>";
            excelFileStr += "<![endif]-->";
            excelFileStr += "</head>";
            excelFileStr += "<body>";
            excelFileStr += excelStr;
            excelFileStr += "</body>";
            excelFileStr += "</html>";
        } else {
            excelFileStr = excelStr;
        }

        if (isIE()) { //如果是IE浏览器  
            try {
                _self.exportFile_IE(fileName, excelFileStr);
            }
            catch (e) {
                alert("导出发生异常:" + e.name + "->" + e.description + "!");
            }
        } else {
            var uri = 'data:text/xls;charset=utf-8,\ufeff' + encodeURIComponent(excelFileStr);
            //创建a标签模拟点击下载
            var downloadLink = document.createElement("a");
            downloadLink.href = uri;
            downloadLink.download = fileName;
            document.body.appendChild(downloadLink);
            downloadLink.click();
            document.body.removeChild(downloadLink);
        }
    },
    exportFile_IE: function (fileName, excelStr) {
        var xlsWin = null;
        if (!!document.all("glbHideFrm")) {
            xlsWin = glbHideFrm;
        } else {
            var width = 1,
                height = 1;
            var openPara = "left=" + (window.screen.width / 2 + width / 2) + ",top=" + (window.screen.height + height / 2) +
                ",scrollbars=no,width=" + width + ",height=" + height;
            xlsWin = window.open("", "_blank", openPara);
        }
        xlsWin.document.write(excelStr);
        xlsWin.document.close();
        xlsWin.document.execCommand('Saveas', true, fileName);
        xlsWin.close();
    },
    getDefaultFileName: function (extension) {
        var d = new Date();
        var curYear = d.getFullYear(),
            curMonth = "" + (d.getMonth() + 1),
            curDate = "" + d.getDate(),
            curHour = "" + d.getHours(),
            curMinute = "" + d.getMinutes(),
            curSecond = "" + d.getSeconds();
        if (curMonth.length == 1) {
            curMonth = "0" + curMonth;
        }
        if (curDate.length == 1) {
            curDate = "0" + curDate;
        }
        if (curHour.length == 1) {
            curHour = "0" + curHour;
        }
        if (curMinute.length == 1) {
            curMinute = "0" + curMinute;
        }
        if (curSecond.length == 1) {
            curSecond = "0" + curSecond;
        }
        var fileName = curYear + curMonth + curDate + curHour + curMinute + curSecond + extension;
        return fileName;
    },
    getFileNameAndExtension: function (fileName) {
        var extension = '.csv',
            _self = exportExcel;
        if (!fileName) {
            fileName = _self.getDefaultFileName(extension);
        } else {
            if (fileName.indexOf('.') < 0) {
                fileName = fileName + extension;
            }
        }
        return fileName;
    }
}


