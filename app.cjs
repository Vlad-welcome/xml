const fs = require("fs");
const path = require("path");
const XlsxTemplate = require("node-xlsx-template");

// 1. Путь к вашему файлу-шаблону
const templatePath = path.join(__dirname, "templates", "template.xlsx");

// 2. Чтение файла шаблона в буфер
const templateData = fs.readFileSync(templatePath);

// 3. Создание экземпляра шаблона из буфера
const template = new XlsxTemplate(templateData);

// 4. Определение данных для подстановки
// Ключи в этом объекте должны точно совпадать с именами в плейсхолдерах ${...}
const values = {
  name: "Ромашка",
};

// 5. Подстановка значений на первый лист (sheet 1)
// Вместо номера можно использовать имя листа, например, 'Sheet1'
template.substitute(1, values);

// 6. Генерация итогового файла в виде бинарного буфера
const outputBuffer = template.generate({ type: "nodebuffer" });

// 7. Сохранение результата в новый файл
const reportPath = path.join(__dirname, "templates", "report.xlsx");
fs.writeFileSync(reportPath, outputBuffer);
console.log("Файл report.xlsx успешно создан!");
