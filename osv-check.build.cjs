const os = require("os");
const path = require("path");
const CFG = process.env.XDG_CONFIG_HOME || path.join(os.homedir(), ".config");
const SKILL = path.join(CFG, "gigatool", "skills", "xlsx");
const ExcelJS = require(path.join(SKILL, "vendor", "exceljs.bundle.cjs"));
const H = require(path.join(SKILL, "helpers", "index.cjs"));

(async () => {
  const wb = new ExcelJS.Workbook();
  const dash = H.addSheet(wb, "Дашборд");
  const osv = H.addSheet(wb, "ОСВ");
  const checks = H.addSheet(wb, "Проверки");

  // ОСВ — демо-данные (фиктивные счета)
  osv.getCell("A3").value = "Счёт";
  osv.getCell("B3").value = "Сальдо нач. Дт, ₽";
  osv.getCell("C3").value = "Сальдо нач. Кт, ₽";
  osv.getCell("D3").value = "Оборот Дт, ₽";
  osv.getCell("E3").value = "Оборот Кт, ₽";
  osv.getCell("F3").value = "Сальдо кон. Дт, ₽";
  osv.getCell("G3").value = "Сальдо кон. Кт, ₽";
  H.headerRow(osv, "A3:G3");

  const rows = [
    ["01", 0, 0, 500000, 0, 500000, 0],
    ["10", 200000, 0, 150000, 120000, 230000, 0],
    ["50", 50000, 0, 30000, 25000, 55000, 0],        // касса
    ["51", 300000, 0, 450000, 420000, 330000, 0],     // расч. счёт
    ["60", 0, 250000, 200000, 300000, 0, 350000],     // счёт поставщиков
    ["62", 150000, 0, 350000, 280000, 220000, 0],     // покупатели
    ["66", 0, 500000, 0, 100000, 0, 600000],          // кредиты
    ["68", 0, 80000, 90000, 100000, 0, 90000],        // налоги
    ["70", 0, 200000, 220000, 240000, 0, 220000],     // оплата труда
    ["80", 0, 1000000, 0, 0, 0, 1000000],             // уставный капитал
  ];
  rows.forEach((r, i) => {
    const row = 4 + i;
    osv.getCell(`A${row}`).value = r[0];
    osv.getCell(`A${row}`).numFmt = H.FMT.text;
    [1, 2, 3, 4, 5, 6].forEach(j => {
      osv.getCell(`${String.fromCharCode(66 + j)}${row}`).value = r[j];
      osv.getCell(`${String.fromCharCode(66 + j)}${row}`).numFmt = H.FMT.rub0;
    });
    // Проверка: у каждого счёта Дт и Кт не м.б. одновременно (по реальной ОСВ они на разных сторонах)
    // Формула: остаток = Дт - Кт на конец (контроль)
    osv.getCell(`H${row}`).value = "Красное сальдо?";
    osv.getCell(`I${row}`).value = { formula: `IF(AND(${H.ref("ОСВ",`$F${row}`)}>0,${H.ref("ОСВ",`$G${row}`)}>0),"ПРОВЕРИТЬ","ок")` };
  });
  osv.getCell("H3").value = "Статус"; osv.getCell("I3").value = "Проверка";
  H.headerRow(osv, "A3:I3");
  H.widths(osv, [["A", 8], ["B", 18], ["C", 18], ["D", 16], ["E", 16], ["F", 18], ["G", 18], ["H", 14], ["I", 16]]);
  H.freeze(osv, 3);

  // Дашборд
  H.titleBand(dash, "A1:D1", "Проверка ОСВ (демо)", "Сбалансированность оборотов и арифметика сальдо.");
  const D = H.ref("ОСВ", "$D$4:$D$200");
  const E = H.ref("ОСВ", "$E$4:$E$200");
  const B = H.ref("ОСВ", "$B$4:$B$200");
  const C = H.ref("ОСВ", "$C$4:$C$200");
  const F = H.ref("ОСВ", "$F$4:$F$200");
  const G = H.ref("ОСВ", "$G$4:$G$200");
  H.kpi(dash, "A3", "Итог оборот Дт, ₽", { formula: `SUM(${D})` }, H.FMT.rub);
  H.kpi(dash, "C3", "Итог оборот Кт, ₽", { formula: `SUM(${E})` }, H.FMT.rub);
  H.kpi(dash, "E3", "Разница Дт−Кт, ₽", { formula: `SUM(${D})-SUM(${E})` }, H.FMT.rub);
  H.kpi(dash, "A5", "Сальдо нач. Дт−Кт", { formula: `SUM(${B})-SUM(${C})` }, H.FMT.rub);
  H.kpi(dash, "C5", "Сальдо кон. Дт−Кт", { formula: `SUM(${F})-SUM(${G})` }, H.FMT.rub);
  dash.getCell("A7").value = "➤ Баланс выполнен, если разницы = 0 и сальдо кон. = сальдо нач. + обороты.";
  H.widths(dash, [["A", 26], ["B", 18], ["C", 18], ["D", 18], ["E", 18]]);

  // Проверки
  checks.getCell("A1").value = "Показатель";
  checks.getCell("B1").value = "Факт";
  checks.getCell("C1").value = "Ожидание";
  checks.getCell("D1").value = "Разница";
  checks.getCell("E1").value = "Допуск";
  checks.getCell("F1").value = "Статус";
  H.headerRow(checks, "A1:F1");
  const checkRows = [
    ["Итог обороты Дт = Кт", `SUM(${D})`, `SUM(${E})`],
    ["Сальдо кон = Дт кон − Кт кон", `SUM(${F})-SUM(${G})`, `SUM(${B})-SUM(${C})+SUM(${D})-SUM(${E})`],
  ];
  checkRows.forEach((r, i) => {
    const row = 2 + i;
    checks.getCell(`A${row}`).value = r[0];
    checks.getCell(`B${row}`).value = { formula: r[1] };
    checks.getCell(`B${row}`).numFmt = H.FMT.rub;
    checks.getCell(`C${row}`).value = { formula: r[2] };
    checks.getCell(`C${row}`).numFmt = H.FMT.rub;
    checks.getCell(`E${row}`).value = 0.01;
    checks.getCell(`D${row}`).value = { formula: `ABS(${H.ref("Проверки",`$B${row}`)}-${H.ref("Проверки",`$C${row}`)})` };
    checks.getCell(`F${row}`).value = { formula: `IF(${H.ref("Проверки",`$D${row}`)}<=${H.ref("Проверки",`$E${row}`)},"OK","FAIL")` };
  });
  H.widths(checks, [["A", 30], ["B", 20], ["C", 20], ["D", 14], ["E", 10], ["F", 10]]);

  await wb.xlsx.writeFile("osv-check-demo.xlsx");
  console.log("wrote osv-check-demo.xlsx");
})();