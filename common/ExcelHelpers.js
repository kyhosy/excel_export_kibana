import ExcelJS from "exceljs"
import {saveAs} from 'file-saver'
import _, {isEmpty, isString} from 'lodash'
import Papa from "papaparse";

export default class ExcelHelpers {
  static read(file) {
    return new Promise(function (resolve) {
      const fileReader = new FileReader();
      fileReader.onload = (e) => {
        const workbook = new ExcelJS.Workbook();
        const buffer = e.target.result;

        return workbook.xlsx.load(buffer).then((wb) => {
          const data = []
          workbook.eachSheet((sheet, id) => {
            sheet.eachRow((row, rowIndex) => {
              data[rowIndex] = row.values
            })
          })

          resolve(data);
        }).catch((error) => {
          console.log("readFile fail", error);
        })
      };
      fileReader.readAsArrayBuffer(file)
    })
  }

  static load(file, callback) {
    const fileReader = new FileReader();
    fileReader.readAsArrayBuffer(file)
    return fileReader.onload = (e) => {
      const workbook = new ExcelJS.Workbook();
      const buffer = e.target.result;

      return workbook.xlsx.load(buffer).then((wb) => {
        if (callback instanceof Function) {
          callback(wb)
        }
      }).catch((error) => {
        console.log("load file fail", error);
      })
    };

  }

  static async saveAsExcel(wb, fileName = 'SavedExcel.xlsx') {
    const buf = await wb.xlsx.writeBuffer()
    saveAs(new Blob([buf]), fileName)
  }

  static autoWidth(ws) {
    ws.columns.forEach(function (column) {
      let maxLength = 0;
      column["eachCell"]({includeEmpty: true}, function (cell) {
        let columnLength = cell.value ? cell.value.toString().length : 10;
        if (columnLength > maxLength) {
          maxLength = columnLength;
        }
      });
      column.width = maxLength < 10 ? 10 : maxLength;
    });
  }

  static async saveAsExcelFromList(items = [], callback, callbackEnd, fileName = 'SavedExcel.xlsx') {
    if (!isEmpty(items)) {
      const wb = new ExcelJS.Workbook()
      const ws = wb.addWorksheet()

      items.forEach((item, index) => {
        if (callback instanceof Function) {
          callback(ws, item, index)
        }
      })

      this.autoWidth(ws)
      if (callbackEnd) {
        callbackEnd(ws)
      }
      this.saveAsExcel(wb, fileName).catch(e => console.log(e))
    }
  }

  static excelColName(n) {
    if (isString(n)) {
      return n;
    }
    let ordA = 'A'.charCodeAt(0);
    let ordZ = 'Z'.charCodeAt(0);
    let len = ordZ - ordA + 1;

    let s = "";
    while (n >= 0) {
      s = String.fromCharCode(n % len + ordA) + s;
      n = Math.floor(n / len) - 1;
    }
    return s;
  }

  static getHeaderCsv(file, onResult){
    console.log('getHeaderCsv>>>>', file)
    Papa.parse(file, {
      header: true,
      preview: 1,
      complete: function(results) {
        console.log(results);
        onResult(results?.meta?.fields)
      }
    });
  }

  static async saveBigFile (headers, rs, fileName, onCompleted, onError ){
    const chunks = _.chunk(rs, 50000);
    console.log('saveBigFile>>>>', chunks.length)
    for (const chunk of chunks) {
      let index = chunks.indexOf(chunk);
      const listData = [
        headers,
        ...chunk
      ]
      console.log('saveBigFile>>>>', index, listData.length)

      await ExcelHelpers.saveAsExcelFromList(listData, (ws, item, id) => {
        ws.addRow(item)
        ws.getRow(1).font = {bold: true}

      },null, `${fileName}_${index+1}.xlsx`).catch(e => {
        console.log('ex>>>>',e)
        if(onError){
          onError(e)
        }
      })
    }

    if(onCompleted){
      onCompleted()
    }
  }

  static downloadExcelFile (file, headers, fileName, onCompleted, onError) {
    // rs.push(headers)
    const rs = []
    Papa.parse(file, {
      // header: true,
      worker: true,
      step: function(row) {
        const item = row.data
        rs.push(item)
      },
      complete: function(results) {
        console.log('results>>>', rs.length, results)
        rs.shift()
        ExcelHelpers.saveBigFile(headers, rs, fileName, onCompleted, onError)
      }
    });
  }

  static downloadExcelFromCsv(file, fileName, mappingColumns, onCompleted, onError){
      ExcelHelpers.getHeaderCsv(file, async (headers) => {
        //TODO: convert key to label
        const trans = headers?.map(col => mappingColumns[col]?.customLabel ?? col)
        ExcelHelpers.downloadExcelFile(file, trans, onCompleted, e => {
          if(onError){
            onError(e)
          }
        })
      })
  }
}
