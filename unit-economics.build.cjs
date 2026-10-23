const os = require("os");
const path = require("path");
const CFG = process.env.XDG_CONFIG_HOME || path.join(os.homedir(), ".config");
const SKILL = path.join(CFG, "gigatool", "skills", "xlsx");
const ExcelJS = require(path.join(SKILL, "vendor", "exceljs.bundle.cjs"));
const H = require(path.join(SKILL, "helpers", "index.cjs"));

(async () => {
  const wb = new ExcelJS.Workbook();
  const dash = H.addSheet(wb, "Дашборд");
  const data = H.addSheet(wb, "SKU");
  const checks = H.addSheet(wb, "Проверки");

  // Демо-данные по SKU (фиктивные, только для примера)
  data.getCell("A3").value = "SKU";
  data.getCell("B3").value = "Себестоим., ₽";
  data.getCell("C3").value = "Комиссия, %";
  data.getCell("D3").value = "Логистика, ₽";
  data.getCell("E3").value = "Эквайринг, %";
  data.getCell("F3").value = "Продано, шт";
  data.getCell("G3").value = "Цена, ₽";
  data.getCell("H3").value = "Маржа/шт, ₽";
  H.headerRow(data, "A3:H3");

  const sku = [
    ["Товар A", 450, 15, 120, 1.5, 120, 1290],
    ["Товар B", 720, 18, 200, 1.5, 80, 2150],
    ["Товар C", 310, 12, 90, 1.5, 200, 890],
    ["Товар D", 980, 20, 260, 1.5, 50, 3200],
  ];
  sku.forEach((r, i) => {
    const row = 4 + i;
    data.getCell(`A${row}`).value = r[0];
    data.getCell(`A${row}`).numFmt = H.FMT.text;
    data.getCell(`B${row}`).value = r[1];
    data.getCell(`B${row}`).numFmt = H.FMT.rub0;
    data.getCell(`C${row}`).value = r[2] / 100;
    data.getCell(`C${row}`).numFmt = H.FMT.pct1;
    data.getCell(`D${row}`).value = r[3];
    data.getCell(`D${row}`).numFmt = H.FMT.rub0;
    data.getCell(`E${row}`).value = r[4] / 100;
    data.getCell(`E${row}`).numFmt = H.FMT.pct1;
    data.getCell(`F${row}`).value = r[5];
    data.getCell(`F${row}`).numFmt = H.FMT.int;
    data.getCell(`G${row}`).value = r[6];
    data.getCell(`G${row}`).numFmt = H.FMT.rub0;
  });

  // Маржа на SKU (формула): цена - себест - комиссия - логистика - эквайринг
  // H = валовая маржа на единицу
  for (let i = 0; i < sku.length; i++) {
    const r = 4 + i;
    data.getCell(`I${r}`).value = null;
    data.getCell(`H${r}`).value = {
      formula: `G${r}-B${r}-G${r}*C${r}-D${r}-G${r}*E${r}`,
      result: Math.round(sku[i][6] - sku[i][1] - sku[i][6] * sku[i][2] / 100 - sku[i][3] - sku[i][6] * sku[i][4] / 100),
    };
    data.getCell(`H${r}`).numFmt = H.FMT.rub0;
  }
  H.widths(data, [["A", 14], ["B", 14], ["C", 14], ["D", 13], ["E", 14], ["F", 12], ["G", 12], ["H", 14]]);
  H.freeze(data, 3);

  // Дашборд
  H.titleBand(dash, "A1:D1", "Юнит-экономика (демо)", "Маржа по каждому SKU для маркетплейса. Входные данные — на листе «SKU».");
  const F = H.ref("SKU", "$A$4:$A$200");
  H.kpi(dash, "A3", "SKU", { formula: `COUNT(${F})` }, H.FMT.int);
  H.kpi(dash, "C3", "Выручка, ₽", { formula: `SUMPRODUCT(${H.ref("SKU","$F$4:$F$200")},${H.ref("SKU","$G$4:$G$200")})` }, H.FMT.rub);
  H.kpi(dash, "E3", "Маржа, ₽", { formula: `SUMPRODUCT(${H.ref("SKU","$F$4:$F$200")},${H.ref("SKU","$H$4:$H$200")})` }, H.FMT.rub);

  dash.getCell("A5").value = "Лучший SKU по марже/шт:";
  dash.getCell("B5").value = { formula: `INDEX(${H.ref("SKU","$A$4:$A$200")},MATCH(MAX(${H.ref("SKU","$H$4:$H$200")}),${H.ref("SKU","$H$4:$H$200")},0))`, result: "Товар D" };
  dash.getCell("B5").numFmt = H.FMT.text;
  dash.getCell("A6").value = "Средняя маржа, %:";
  dash.getCell("B6").value = { formula: `SUMPRODUCT(${H.ref("SKU","$F$4:$F$200")},${H.ref("SKU","$H$4:$H$200")})/SUMPRODUCT(${H.ref("SKU","$F$4:$F$200")},${H.ref("SKU","$G$4:$G$200")})` };
  dash.getCell("B6").numFmt = H.FMT.pct1;
  H.widths(dash, [["A", 26], ["B", 16], ["C", 18], ["D", 16], ["E", 14]]);

  // Проверки
  checks.getCell("A1").value = "Показатель";
  checks.getCell("B1").value = "Факт";
  checks.getCell("C1").value = "Ожидание";
  checks.getCell("D1").value = "Разница";
  checks.getCell("E1").value = "Допуск";
  checks.getCell("F1").value = "Статус";
  H.headerRow(checks, "A1:F1");
  const Fv = H.ref("SKU", "$G$4:$G$200");
  const Fa = H.ref("SKU", "$F$4:$F$200");
  checks.getCell("A2").value = "Выручка (контроль)";
  checks.getCell("B2").value = { formula: `SUMPRODUCT(${Fa},${Fv})`, result: 120 * 1290 + 80 * 2150 + 200 * 890 + 50 * 3200 };
  checks.getCell("C2").value = 120 * 1290 + 80 * 2150 + 200 * 890 + 50 * 3200;
  checks.getCell("E2").value = 1;
  checks.getCell("F2").value = { formula: `IF(ABS(${H.ref("Проверки","$B$2")}-${H.ref("Проверки","$C$2")})<=${H.ref("Проверки","$E$2")},"OK","FAIL")` };
  checks.getCell("D2").value = { formula: `ABS(${H.ref("Проверки","$B$2")}-${H.ref("Проверки","$C$2")})` };
  H.widths(checks, [["A", 26], ["B", 16], ["C", 16], ["D", 12], ["E", 10], ["F", 10]]);

  await wb.xlsx.writeFile("unit-economics-demo.xlsx");
  console.log("wrote unit-economics-demo.xlsx");
})();