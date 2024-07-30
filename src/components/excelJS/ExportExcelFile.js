import axios from "axios";
import * as ExcelJS from "exceljs";

const toDataURL = (url) => {
  const promise = new Promise((resolve, reject) => {
    var xhr = new XMLHttpRequest();
    xhr.onload = function () {
      var reader = new FileReader();
      reader.readAsDataURL(xhr.response);
      reader.onloadend = function () {
        resolve({ base64Url: reader.result });
      };
    };
    xhr.open("GET", url);
    xhr.responseType = "blob";
    xhr.send();
  });

  return promise;
};

export const exportExcelFile = async (data) => {
  const workbook = new ExcelJS.Workbook();
  // create new sheet with pageSetup settings for A4 - landscape
  const worksheet = workbook.addWorksheet('sheet', {
    pageSetup: { paperSize: 9, orientation: 'landscape' },
    headerFooter: { firstHeader: "Header ECPay", firstFooter: "Footer ECPay" }
  });

  // Image

  const imageBuffer = await axios.get('/striker.png', { responseType: 'arraybuffer' });

  const StrikerLogo = workbook.addImage({
    buffer: imageBuffer.data,
    extension: 'png',
  });

  worksheet.addImage(StrikerLogo, 'C2:E6');

  worksheet.mergeCells('C8:E8');

  worksheet.getCell('B8').value = "REPORTE ECPAY 2";
  worksheet.getCell('E8').value = " ";
  worksheet.getCell('B8').font = { bold: true, size: 24, alignment: { horizontal: "center", vertical: "middle" } };
  worksheet.getCell('B8').alignment = { horizontal: "center", vertical: "middle" };

  worksheet.addRow([]);

  const cellTemp = worksheet.addRow(["Nombre", data[0].fullName, , "Documento", data[0].document]);

  cellTemp.eachCell((cell, colNumber) => {
    cell.style = { alignment: { horizontal: "left" } }

    cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
    if (colNumber == 1 || colNumber == 4) {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0074FF' } };
      cell.font = { bold: true };
    }
  });

  const cellTemp2 = worksheet.addRow(["Id del empleado", data[0].employeeId, , "Descripción", data[0].description]);

  cellTemp2.eachCell((cell, colNumber) => {
    cell.style = { alignment: { horizontal: "left" } }

    cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
    if (colNumber == 1 || colNumber == 4) {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'AA0074FF' } };
      cell.font = { bold: true };
    }
  });

  const cellTemp3 = worksheet.addRow(["Fecha inicial del contrato", new Date(data[0].initialDateContract), , "Fecha final del contrato", new Date(data[0].finalDateContract)]);

  cellTemp3.eachCell((cell, colNumber) => {
    cell.style = { alignment: { horizontal: "left" } }

    cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
    if (colNumber == 1 || colNumber == 4) {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '660074FF' } };
      cell.font = { bold: true };
    }
  });

  const cellTemp4 = worksheet.addRow(["Tipo de contrato", data[0].contractType, , "Tipo de nómina", data[0].payrollType]);

  cellTemp4.eachCell((cell, colNumber) => {
    cell.style = { alignment: { horizontal: "left" } }
    cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
    if (colNumber == 1 || colNumber == 4) {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '000074FF' } };
      cell.font = { bold: true };
    }
  });

  worksheet.addRow();

  //----------------------TABLE--------------------------------//
  const tableHeader = worksheet.addRow(columnNames);

  tableHeader.eachCell((cell, colNumber) => {
    cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0074FF' } };
    cell.font = { bold: true };
    worksheet.getColumn(colNumber).width = 25;
  });


  // Assign only data on columns mapped from object
  let partialData = data.map(
    (obj) => columnNames.map(element => obj[element])
  );

  partialData.forEach(item => {
    const rowData = Object.values(item);
    const row = worksheet.addRow(rowData);
    row.eachCell((cell, colNumber) => {
      cell.style = {
        numFmt: "$###,###,###,###.00",
        alignment: { horizontal: "center" }
      }
      cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
      if (colNumber >= 2 && colNumber <= 4) {
        cell.alignment = { horizontal: 'center' };
      }
    });
  });

  worksheet.columns.forEach(function (column, i) {
    let maxLength = 0;
    column["eachCell"]({ includeEmpty: true }, function (cell) {
      var columnLength = cell.value ? cell.value.toString().length : 10;
      if (columnLength > maxLength) {
        maxLength = columnLength;
      }
    });
    column.width = maxLength < 10 ? 10 : maxLength;
  });

  // Move everything one cell to the right
  worksheet.spliceColumns(1, 0, [])

  workbook.xlsx.writeBuffer().then(function (data) {
    const blob = new Blob([data], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const url = window.URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = "download.xlsx";
    anchor.click();
    window.URL.revokeObjectURL(url);
  });
};

const columnNames = [
  "payType",
  "nature",
  "conceptFull",
  "amount",
  "value"
]
