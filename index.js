const xlsx = require("xlsx-style");
const glob = require("glob");

xlsx.writeFile(
  glob
    .sync("from/**/*.xlsx")
    .map((file) =>
      xlsx.readFile(file, { cellStyles: true, cellNF: true, cellDates: true })
    )
    .map((workbook) => {
      console.log(workbook);
      return workbook;
    })
    .map((workbook) => ({
      ...workbook,
      sheet: workbook.Sheets[workbook.SheetNames[0]],
      sheetName: workbook.SheetNames[0],
    }))
    .reduce(
      (obj, item) => {
        obj.Sheets[item.sheetName] = item.sheet;
        obj.SheetNames.push(item.sheetName);
        return obj;
      },
      {
        SheetNames: [],
        Sheets: {},
        Props: {},
      }
    ),
  "to.xlsx"
);
