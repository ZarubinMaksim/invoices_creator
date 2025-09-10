const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const xlsx = require('xlsx');
const app = express();
const PORT = 4000;
const puppeteer = require('puppeteer');

// Папка для сохранённых PDF
const pdfFolder = path.join(__dirname, 'saved_pdf');
if (!fs.existsSync(pdfFolder)) fs.mkdirSync(pdfFolder);

// Префикс маршрута
const ROUTE_PREFIX = '/invoices';

// Папка для загрузки файлов
const uploadFolder = path.join(__dirname, 'uploads');
if (!fs.existsSync(uploadFolder)) fs.mkdirSync(uploadFolder);

// Настройка multer с лимитом 100 MB
const storage = multer.diskStorage({
    destination: (req, file, cb) => cb(null, uploadFolder),
    filename: (req, file, cb) => {
        const ext = path.extname(file.originalname);
        const name = path.basename(file.originalname, ext);
        cb(null, `${name}-${Date.now()}${ext}`);
    }
});
const upload = multer({
    storage,
    limits: { fileSize: 100 * 1024 * 1024 } // 100 MB
});

// Главная страница с формой загрузки
app.get(`${ROUTE_PREFIX}/`, (req, res) => {
    res.send(`
        <h1>Загрузка Excel файла</h1>
        <form action="${ROUTE_PREFIX}/upload" method="post" enctype="multipart/form-data">
            <input type="file" name="excel" accept=".xls,.xlsx" required />
            <button type="submit">Загрузить</button>
        </form>
    `);
});

// Маршрут для загрузки файла
app.post(`${ROUTE_PREFIX}/upload`, upload.single('excel'), async (req, res) => {
  if (!req.file) return res.status(400).send('Файл не загружен');

  const workbook = xlsx.readFile(req.file.path);
  const sheetIndex = workbook.SheetNames.length - 4; // предпоследний лист
  const sheetName = workbook.SheetNames[sheetIndex];
  const worksheet = workbook.Sheets[sheetName];

  const data = xlsx.utils.sheet_to_json(worksheet, { defval: '' });

  let htmlOverview = `<h1>Файл успешно загружен: ${req.file.filename}</h1>`;
  htmlOverview += `<h2>Созданные PDF:</h2><ul>`;

  for (let rowIndex = 2; rowIndex < data.length; rowIndex++) { // начиная с третьей строки
    const row = data[rowIndex];
    const name = row['Guest name'] || '';
    const room = row['Room no.'] || '';
    const amount = row['Total amount'] || '';

    // Подставляем данные в HTML шаблон
    let invoiceHtml = fs.readFileSync(path.join(__dirname, 'invoice_template.html'), 'utf-8');
    invoiceHtml = invoiceHtml.replace('{{name}}', name)
                             .replace('{{room}}', room)
                             .replace('{{amount}}', amount);

    // Генерируем PDF
    const browser = await puppeteer.launch({ args: ['--no-sandbox'], executablePath: '/usr/bin/chromium-browser', headless: true });
    const page = await browser.newPage();
    await page.setContent(invoiceHtml, { waitUntil: 'networkidle0' });

    const pdfPath = path.join(pdfFolder, `${name.replace(/\s+/g, '_')}_${room}_${Date.now()}.pdf`);
    await page.pdf({ path: pdfPath, format: 'A4', printBackground: true });
    await browser.close();

    htmlOverview += `<li>${pdfPath}</li>`;
  }

  htmlOverview += '</ul>';
  res.send(htmlOverview);
});


// Слушаем все внешние подключения
app.listen(PORT, '0.0.0.0', () => {
    console.log(`Invoices server запущен на http://38.244.150.204:${PORT}${ROUTE_PREFIX}`);
});
