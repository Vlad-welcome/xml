const ExcelJS = require('exceljs');

async function fillTableTemplate(templatePath, outputPath) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(templatePath);
  const sheet = workbook.getWorksheet(1);

  // JSON с данными для таблицы
  const items = [
    { name: 'Ноутбук',    quantity: 2, price: 1200 },
    { name: 'Монитор',    quantity: 3, price: 300  },
    { name: 'Клавиатура', quantity: 5, price: 50   },
    { name: 'Мышь',       quantity: 10, price: 25  },
  ];

  // Строка 2 — образец стиля. Начинаем вставку с позиции 2.
  // Метод insertRow(position, values, inheritFrom) копирует стиль со сдвигаемой строки.
  const startRow = 2;

  items.forEach((item, index) => {
    const rowNumber = startRow + index;

    // 'i+' — наследовать стиль от строки выше (включая пустые ячейки)
    // При первой вставке стиль берётся со строки-образца, при последующих — с уже вставленной.
    const row = sheet.insertRow(rowNumber, [
      item.name,
      item.quantity,
      item.price,
      item.quantity * item.price, // вычисляем сумму в JS
    ], 'i+');

    // Дополнительно можно задать формат ячеек
    row.getCell(3).numFmt = '#,##0.00';
    row.getCell(4).numFmt = '#,##0.00';
  });

  // Удаляем строку-образец, если она больше не нужна
  sheet.spliceRows(startRow + items.length, 1);

  await workbook.xlsx.writeFile(outputPath);
  console.log('Таблица заполнена:', outputPath);
}

fillTableTemplate('./template.xlsx', './filled.xlsx');