const ExcelJS = require("exceljs");
const path = require("path");

module.exports = async function () {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Me";
  workbook.lastModifiedBy = "Her";
  workbook.created = new Date(1985, 8, 30);
  workbook.modified = new Date();
  workbook.lastPrinted = new Date(2016, 9, 27);

  const worksheet = workbook.addWorksheet('New Sheet');

  worksheet.columns = [
    { header: 'Id', key: 'id' },
    { header: 'Name', key: 'name' },
    { header: 'Age', key: 'age' }
  ];

  const rows = [
    [3,'Alex','44'],
    {id:4, name: 'Margaret', age: 32}
  ];

  worksheet.addRows(rows);

  const imageId1 = workbook.addImage({
    filename: 'path/to/image.jpg',
    extension: 'jpeg',
  });

  workbook.xlsx.writeBuffer().then((data) => {
    const blob = new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=utf-8' });
    saveAs(blob, 'test.xlsx');
  });
};
