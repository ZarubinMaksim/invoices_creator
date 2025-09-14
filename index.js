const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const xlsx = require('xlsx');
const app = express();
const PORT = 4000;
const puppeteer = require('puppeteer');
const { execSync } = require('child_process');
const toThaiBahtText = require('thai-baht-text');
const { toWords } = require('number-to-words');
const archiver = require('archiver');
const nodemailer = require('nodemailer');
require('dotenv').config();

console.log('🚀 Инициализация сервера...');

// Убиваем все висящие процессы Chromium перед запуском
console.log('🔄 Убиваем висящие процессы Chromium...');
try {
    execSync('pkill -f chromium', { stdio: 'ignore' });
    console.log('✅ Процессы Chromium завершены');
} catch (error) {
    console.log('ℹ️ Не было процессов Chromium для завершения');
}

// Папка для сохранённых PDF
const pdfFolder = path.join(__dirname, 'saved_pdf');
if (!fs.existsSync(pdfFolder)) {
    console.log('📁 Создаем папку для PDF:', pdfFolder);
    fs.mkdirSync(pdfFolder, { recursive: true });
} else {
    console.log('📁 Папка для PDF уже существует:', pdfFolder);
}



// Префикс маршрута
const ROUTE_PREFIX = '/invoices';

// Делаем папку доступной по URL
app.use(`${ROUTE_PREFIX}/pdf`, express.static(pdfFolder));

// Папка для загрузки файлов
const uploadFolder = path.join(__dirname, 'uploads');
if (!fs.existsSync(uploadFolder)) {
    console.log('📁 Создаем папку для загрузок:', uploadFolder);
    fs.mkdirSync(uploadFolder, { recursive: true });
} else {
    console.log('📁 Папка для загрузок уже существует:', uploadFolder);
}

// Настройка multer
const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        console.log('📂 Сохраняем файл в:', uploadFolder);
        cb(null, uploadFolder);
    },
    filename: (req, file, cb) => {
        const ext = path.extname(file.originalname);
        const name = path.basename(file.originalname, ext);
        const filename = `${name}-${Date.now()}${ext}`;
        console.log('📝 Новое имя файла:', filename);
        cb(null, filename);
    }
});
const upload = multer({
    storage,
    limits: { fileSize: 100 * 1024 * 1024 }
});

// Глобальная переменная для браузера
let browserInstance = null;

// Функция для получения экземпляра браузера
async function getBrowser() {
    if (!browserInstance) {
        console.log('🌐 Запускаем браузер...');
        
        const launchOptions = {
            args: [
                '--no-sandbox',
                '--disable-setuid-sandbox',
                '--disable-dev-shm-usage',
                '--disable-gpu',
                '--single-process',
                '--no-zygote',
                '--disable-extensions',
                '--disable-software-rasterizer',
                '--disable-background-timer-throttling',
                '--disable-backgrounding-occluded-windows',
                '--disable-renderer-backgrounding'
            ],
            headless: 'new',
            timeout: 120000,
            executablePath: '/snap/chromium/current/usr/lib/chromium-browser/chrome'
        };
        
        console.log('⚙️ Параметры запуска:', launchOptions);
        
        try {
            browserInstance = await puppeteer.launch(launchOptions);
            console.log('✅ Браузер успешно запущен');
            
            // Проверяем версию
            const version = await browserInstance.version();
            console.log('🌐 Версия браузера:', version);
            
        } catch (error) {
            console.error('❌ Ошибка запуска браузера:', error);
            
            // Пробуем альтернативный путь
            console.log('🔄 Пробуем альтернативный путь...');
            launchOptions.executablePath = '/usr/bin/chromium-browser';
            
            try {
                browserInstance = await puppeteer.launch(launchOptions);
                console.log('✅ Браузер запущен с альтернативным путем');
            } catch (retryError) {
                console.error('❌ Ошибка при повторной попытке запуска:', retryError);
                throw retryError;
            }
        }
    }
    
    return browserInstance;
}

// Главная страница с формой загрузки
app.get(`${ROUTE_PREFIX}/`, (req, res) => {
    console.log('📋 Получен GET запрос на главную страницу');
    res.send(`
        <h1>Загрузка Excel файла</h1>
        <form action="${ROUTE_PREFIX}/upload" method="post" enctype="multipart/form-data">
            <input type="file" name="excel" accept=".xls,.xlsx" required />
            <button type="submit">Загрузить</button>
        </form>
    `);
});

// Маршрут для скачивания всех PDF в ZIP
app.get(`${ROUTE_PREFIX}/download-all`, (req, res) => {
  const zipName = `all_invoices_${Date.now()}.zip`;
  res.setHeader('Content-Disposition', `attachment; filename=${zipName}`);
  res.setHeader('Content-Type', 'application/zip');

  const archive = archiver('zip', { zlib: { level: 9 } });

  archive.on('error', err => {
      console.error('❌ Ошибка архивации:', err);
      res.status(500).send({ error: err.message });
  });

  // Прямо в поток ответа
  archive.pipe(res);

  // Добавляем все PDF из папки
  fs.readdirSync(pdfFolder).forEach(file => {
      const filePath = path.join(pdfFolder, file);
      archive.file(filePath, { name: file });
  });

  archive.finalize();
});

// Транспорт для отправки Gmail (нужен app password)
const transporter = nodemailer.createTransport({
  service: 'gmail',
  auth: {
      user: process.env.GMAIL_USER,
      pass: process.env.GMAIL_PASS  // не обычный пароль, а пароль приложения Google
  }
});


// API для отправки писем
// API для отправки писем
app.post(`${ROUTE_PREFIX}/send-emails`, express.json(), async (req, res) => {
  const rows = req.body.rows || [];
  const results = [];

  for (const row of rows) {
    try {
      await transporter.sendMail({
        from: '"Invoices" <gsm@lagreenhotel.com>',
        to: row.email,
        subject: `Ваш счёт за номер ${row.room} в La Green Hotel & Residence`,
        text: `Здравствуйте, ${row.name}! Во вложении ваш счет за номер ${row.room}.`,
        attachments: [
          {
            filename: path.basename(row.pdf),
            path: path.join(__dirname, row.pdf.replace(`${ROUTE_PREFIX}/pdf/`, 'saved_pdf/'))
          }
        ]
      });
      results.push({ room: row.room, name: row.name, email: row.email, status: 'Отправлено' });
    } catch (err) {
      console.error('Ошибка отправки на', row.email, err);
      results.push({ room: row.room, name: row.name, email: row.email, status: 'Ошибка' });
    }
  }

  res.json({ results });
});


function getCurrentDate() {
  const today = new Date();
  const day = String(today.getDate()).padStart(2, '0');
  const month = String(today.getMonth() + 1).padStart(2, '0');
  const year = today.getFullYear();
  return `${day}/${month}/${year}`;
}

function excelDateToDDMMYYYY(serial) {
  const excelEpoch = new Date(Date.UTC(1899, 11, 30)); // база для Excel
  const days = Math.floor(serial);
  const milliseconds = days * 24 * 60 * 60 * 1000;
  const date = new Date(excelEpoch.getTime() + milliseconds);

  const dd = String(date.getUTCDate()).padStart(2, '0');
  const mm = String(date.getUTCMonth() + 1).padStart(2, '0'); // месяцы с 0
  const yyyy = date.getUTCFullYear();

  return `${dd}/${mm}/${yyyy}`;
}

function generateInvoiceNumber(counter, serial) {
  const excelEpoch = new Date(Date.UTC(1899, 11, 30)); // база для Excel
  const days = Math.floor(serial);
  const milliseconds = days * 24 * 60 * 60 * 1000;
  const date = new Date(excelEpoch.getTime() + milliseconds);

  const mm = String(date.getUTCMonth() + 1).padStart(2, '0'); // месяцы с 0
  const yyyy = date.getUTCFullYear();

  const number = String(counter).padStart(3, '0'); // порядковый номер с ведущими нулями
  return `PS${yyyy}${mm}-${number}`;
}

let logQueue = [];

app.get(`${ROUTE_PREFIX}/logs`, (req, res) => {
  res.setHeader('Content-Type', 'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache');
  res.setHeader('Connection', 'keep-alive');

  // Отправка логов по очереди каждые 200 мс
  const interval = setInterval(() => {
    while (logQueue.length > 0) {
      const msg = logQueue.shift();
      res.write(`data: ${msg}\n\n`);
    }
  }, 200);

  req.on('close', () => {
    clearInterval(interval);
  });
});

// Функция для логирования и добавления сообщений в очередь
function logToBrowser(msg) {
  console.log(msg); // обычный консоль лог
  logQueue.push(msg);
}

// Маршрут для загрузки файла
app.post(`${ROUTE_PREFIX}/upload`, upload.single('excel'), async (req, res) => {
    console.log('📤 Получен POST запрос на загрузку файла');
    logToBrowser('📤 Получен POST запрос на загрузку файла')
    if (!req.file) {
        console.log('❌ Файл не загружен');
        logToBrowser('❌ Файл не загружен')
        return res.status(400).send('Файл не загружен');
    }

    console.log('✅ Файл загружен:', req.file.filename);
    logToBrowser('✅ Файл загружен:', req.file.filename)

    try {
        console.log('📖 Читаем Excel файл...');
        logToBrowser('📖 Читаем Excel файл...');

        const workbook = xlsx.readFile(req.file.path);
        console.log('✅ Файл прочитан успешно');
        logToBrowser('✅ Файл прочитан успешно');

        const sheetIndex = workbook.SheetNames.length - 4;
        const sheetName = workbook.SheetNames[sheetIndex];
        console.log('📑 Выбран лист:', sheetName);
        logToBrowser('📑 Выбран лист:', sheetName)
        
        const worksheet = workbook.Sheets[sheetName];
        const data = xlsx.utils.sheet_to_json(worksheet, { defval: '' });
        
        console.log('📈 Найдено строк:', data.length);
        logToBrowser('📈 Найдено строк:', data.length)

        res.writeHead(200, {
            'Content-Type': 'text/html; charset=utf-8',
            'Transfer-Encoding': 'chunked'
        });
        
        res.write(`
        <h1>Файл успешно загружен: ${req.file.filename}</h1>
        <h2>Создание PDF:</h2>
        
        <!-- Таблица с PDF -->
        <table id="pdf-table" border="1" cellspacing="0" cellpadding="5" style="border-collapse: collapse; width: 100%;">
          <thead>
            <tr style="background-color: #f2f2f2;">
              <th>№</th>
              <th>Комната</th>
              <th>Имя</th>
              <th>Почта</th>
              <th><input type="checkbox" id="select-all" /> Все</th>
              <th>Вода</th>
              <th>Свет</th>
              <th>Всего</th>
              <th>Статус</th>
              <th>Счёт</th>
            </tr>
          </thead>
          <tbody></tbody>
        </table>
        
        <!-- Кнопки -->
        <button onclick="window.location.href='${ROUTE_PREFIX}/download-all'" 
          style="margin-top:20px; padding:10px 20px; background:#4CAF50; color:white; border:none; border-radius:5px;">
          Скачать все счета ZIP
        </button>
        
        <button onclick="sendSelectedEmails()" 
          style="margin-top:20px; margin-left:20px; padding:10px 20px; background:#2196F3; color:white; border:none; border-radius:5px;">
          Отправить выбранные счета на почту
        </button>
        
        <!-- Блок логов -->
        <h2>Логи обработки</h2>
        <div id="server-logs" style="border:1px solid #ccc; padding:10px; height:200px; overflow-y:auto; margin-top:10px;">
          <strong>Логи сервера:</strong><br>
        </div>
        
        <h2>Результаты рассылки</h2>
        <table id="email-results" border="1" cellspacing="0" cellpadding="5" style="border-collapse: collapse; width: 100%;">
          <thead>
            <tr style="background-color: #f2f2f2;">
              <th>№</th>
              <th>Комната</th>
              <th>ФИО</th>
              <th>Почта</th>
              <th>Статус</th>
            </tr>
          </thead>
          <tbody></tbody>
        </table>
        
        <div id="email-status" style="margin-top:20px; font-weight:bold;"></div>
        
        <script>
        let counter = 0;
        function addPdfRow(room, name, email, water, electricity, total, status, pdfPath) {
          counter++;
          const tbody = document.querySelector('#pdf-table tbody');
          const row = document.createElement('tr');
        
          let statusCell = '<td style="background:' + (status === 'success' ? '#c6efce' : '#ffc7ce') +
                           '; text-align:center; font-weight:bold;">' +
                           (status === 'success' ? 'SUCCESS' : 'ERROR') + '</td>';
        
          let downloadCell = '';
          if (status === 'success') {
            downloadCell = '<td><a href="' + pdfPath + '" target="_blank" ' +
                           'style="display:inline-block; padding:5px 10px; background:#4CAF50; color:white; text-decoration:none; border-radius:5px;">Скачать</a></td>';
          } else {
            downloadCell = '<td>-</td>';
          }
        
          row.innerHTML = '<td>' + counter + '</td>' +
                          '<td>' + room + '</td>' +
                          '<td>' + name + '</td>' +
                          '<td>' + email + '</td>' +
                          '<td><input type="checkbox" class="email-checkbox" ' +
        
        `);
        

        // Получаем браузер
        console.log('🖥️ Получаем экземпляр браузера...');
        const browser = await getBrowser();
        console.log('✅ Браузер готов к работе');
        
        let successCount = 0;
        let errorCount = 0;
        let invoiceCount = 0

        for (let rowIndex = 2; rowIndex < data.length; rowIndex++) {
            invoiceCount += 1
            const row = data[rowIndex];
            const name = row['Guest name'] || '';
            const room = row['Room no.'] || '';
            //const rawEmail = row['Guest e-mail'] || ''; //удалить когда колонки емаил и тел будут отдельные
            //const email = rawEmail.split(/[\s/]/)[0].trim();     //удалить когда колонки емаил и тел будут отдельные        
            const email = '89940028777@ya.ru'
            const water_start = (parseFloat(row['Water Meter numbers']) || 0).toFixed(2);
            const water_end = (parseFloat(row['__EMPTY_2']) || 0).toFixed(2);
            const water_consumption = (parseFloat(row['Water consumption']) || 0).toFixed(2);
            const water_price = 89;
            const water_total = (parseFloat(row['__EMPTY_3']) || 0).toFixed(2);
            const electricity_start = (parseFloat(row['Electricity Meter numbers']) || 0).toFixed(2);
            const electricity_end = (parseFloat(row['__EMPTY_4']) || 0).toFixed(2);
            const electricity_consumption = (parseFloat(row['Eletricity']) || 0).toFixed(2);
            const electricity_price = 8;
            const electricity_total = (parseFloat(row['__EMPTY_5']) || 0).toFixed(2);
            const amount_total = (parseFloat(row['Before amount']) || 0).toFixed(2);
            const amount_before_vat = (parseFloat(row['Before amount']) || 0).toFixed(2);
            const vat = (parseFloat(row['SVC']) || 0).toFixed(2);
            const amount_total_net = (parseFloat(row['Total amount']) || 0).toFixed(2);
            const invoice_number = generateInvoiceNumber(invoiceCount, row['Period Check']); 
            const date_from = excelDateToDDMMYYYY(row['Period Check']) || '';
            const date_to = excelDateToDDMMYYYY(row['__EMPTY_1']) || '';
            const date_of_creating = getCurrentDate()
            const total_in_thai = toThaiBahtText(amount_total_net)
            const total_in_english = toWords(amount_total_net)



            console.log(`📊 Обрабатываем строку ${rowIndex}:`, { 
              name, 
              room, 
              water_start, 
              water_end, 
              water_consumption, 
              water_price, 
              water_total, 
              electricity_start, 
              electricity_end, 
              electricity_consumption, 
              electricity_price, 
              electricity_total, 
              amount_total, 
              amount_before_vat, 
              vat, 
              amount_total_net,
              invoice_number,
            date_from,
          date_to,
          date_of_creating,
          total_in_thai,
        total_in_english });

            if (!name && !room) {
                console.log('⏭️ Пропускаем пустую строку');
                continue;
            }

            try {
                console.log('📄 Читаем HTML шаблон...');
                const logoPath = path.join(__dirname, 'img/logo.png');
                const qrPath = path.join(__dirname, 'img/qr.png');
                const logoBase64 = fs.readFileSync(logoPath).toString('base64');
                const qrBase64 = fs.readFileSync(qrPath).toString('base64');
                const logoDataUri = `data:image/png;base64,${logoBase64}`;
                const qrDataUri = `data:image/png;base64,${qrBase64}`;
                let invoiceHtml = fs.readFileSync(path.join(__dirname, 'invoice_template.html'), 'utf-8');
                invoiceHtml = invoiceHtml.replace('{{name}}', name)
                                         .replace('{{room}}', room)
                                         .replace('{{water_start}}', water_start)
                                         .replace('{{water_end}}', water_end)
                                         .replace('{{water_consumption}}', water_consumption)
                                         .replace('{{water_price}}', water_price)
                                         .replace('{{water_total}}', water_total)
                                         .replace('{{electricity_start}}', electricity_start)
                                         .replace('{{electricity_end}}', electricity_end)
                                         .replace('{{electricity_consumption}}', electricity_consumption)
                                         .replace('{{electricity_price}}', electricity_price)
                                         .replace('{{electricity_total}}', electricity_total)
                                         .replace('{{amount_total}}', amount_total)
                                         .replace('{{amount_before_vat}}', amount_before_vat)
                                         .replace('{{vat}}', vat)
                                         .replace('{{amount_total_net}}', amount_total_net)
                                         .replace('{{invoice_number}}', invoice_number)
                                         .replace('{{date_from}}', date_from)
                                         .replace('{{date_to}}', date_to)
                                         .replace('{{date_of_creating}}', date_of_creating)
                                         .replace('{{total_in_thai}}', total_in_thai)
                                         .replace('{{total_in_english}}', total_in_english)
                                         .replace('{{qr_base64}}', qrDataUri)
                                         .replace('{{logo_base64}}', logoDataUri);

                // Создаем новую страницу
                console.log('🆕 Создаем новую страницу...');
                const page = await browser.newPage();
                
                console.log('🔄 Устанавливаем контент...');
                await page.setContent(invoiceHtml, { 
                    waitUntil: 'networkidle0',
                    timeout: 30000
                });
                const pdfFileName = `${name.replace(/\s+/g, '_')}_${room}_${Date.now()}.pdf`;
                const pdfPath = path.join(pdfFolder, pdfFileName);
                console.log('🖨️ Генерируем PDF:', pdfPath);
                
                await page.pdf({ 
                    path: pdfPath, 
                    format: 'A4', 
                    printBackground: true,
                    timeout: 30000
                });
                
                console.log('✅ PDF успешно создан');
                await page.close();
                const pdfUrl = `${ROUTE_PREFIX}/pdf/${pdfFileName}`;

                res.write(`<script>addPdfRow("${room}", "${name}", "${email}", "${water_total}", "${electricity_total}", "${amount_total}", "success", "${pdfUrl}");</script>`);
                successCount++;
                
            } catch (error) {
                console.error('❌ Ошибка:', error);
                errorCount++;
                res.write(`<script>addPdfRow("${room}", "${name}", "${email}", "${water_total}", "${electricity_total}", "${amount_total}", "error", "");</script>`);
              }
        }

        res.write(`<h3>Генерация завершена! Успешно: ${successCount}, Ошибок: ${errorCount}</h3>`);
        res.end();
        console.log(`✅ Обработка завершена. Успешно: ${successCount}, Ошибок: ${errorCount}`);

    } catch (error) {
        console.error('❌ Критическая ошибка:', error);
        res.status(500).send('Ошибка: ' + error.message);
    }
});

// Обработка завершения приложения
process.on('SIGINT', async () => {
    console.log('\n🛑 Получен сигнал SIGINT, завершаем работу...');
    if (browserInstance) {
        console.log('❌ Закрываем браузер...');
        await browserInstance.close();
        console.log('✅ Браузер закрыт');
    }
    console.log('👋 Завершение работы');
    process.exit();
});

// Слушаем все внешние подключения
app.listen(PORT, '0.0.0.0', () => {
    console.log(`✅ Invoices server запущен на порту ${PORT}`);
    console.log(`📋 Доступно по: http://38.244.150.204:${PORT}${ROUTE_PREFIX}`);
});