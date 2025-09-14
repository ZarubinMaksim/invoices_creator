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

app.use(express.json());

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
            
            const version = await browserInstance.version();
            console.log('🌐 Версия браузера:', version);
            
        } catch (error) {
            console.error('❌ Ошибка запуска браузера:', error);
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

// Настройка nodemailer
const transporter = nodemailer.createTransport({
  service: 'gmail',
  auth: {
      user: 'gsm@lagreenhotel.com',
      pass: 'ljhg ojga fxoj dltc'  // лучше использовать app password
  }
});

async function sendInvoiceEmail(toEmail, pdfPaths) {
    try {
        const attachments = pdfPaths.map(pdfPath => ({
            filename: path.basename(pdfPath),
            path: pdfPath
        }));

        const mailOptions = {
            from: `"Компания" <gsm@lagreenhotel.com>`,
            to: toEmail,
            subject: 'Ваши счета',
            text: 'Здравствуйте! Во вложении ваши счета.',
            attachments
        };

        const info = await transporter.sendMail(mailOptions);
        console.log('📧 Письмо отправлено:', info.response);
        return true;
    } catch (error) {
        console.error('❌ Ошибка отправки письма:', error);
        return false;
    }
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

  archive.pipe(res);

  fs.readdirSync(pdfFolder).forEach(file => {
      const filePath = path.join(pdfFolder, file);
      archive.file(filePath, { name: file });
  });

  archive.finalize();
});

// Маршрут для отправки выбранных PDF по почте
app.post(`${ROUTE_PREFIX}/send-selected`, async (req, res) => {
    const { email, pdfs } = req.body;
    if (!email || !pdfs || !Array.isArray(pdfs) || pdfs.length === 0) {
        return res.status(400).json({ error: 'Не указаны email или PDF файлы' });
    }

    const pdfPaths = pdfs.map(p => path.join(pdfFolder, path.basename(p)));
    const success = await sendInvoiceEmail(email, pdfPaths);
    res.json({ success });
});

// Маршрут для загрузки Excel и генерации PDF
app.post(`${ROUTE_PREFIX}/upload`, upload.single('excel'), async (req, res) => {
    console.log('📤 Получен POST запрос на загрузку файла');
    
    if (!req.file) {
        console.log('❌ Файл не загружен');
        return res.status(400).send('Файл не загружен');
    }

    console.log('✅ Файл загружен:', req.file.filename);

    try {
        console.log('📖 Читаем Excel файл...');
        const workbook = xlsx.readFile(req.file.path);
        console.log('✅ Файл прочитан успешно');
        
        const sheetIndex = workbook.SheetNames.length - 4;
        const sheetName = workbook.SheetNames[sheetIndex];
        console.log('📑 Выбран лист:', sheetName);
        
        const worksheet = workbook.Sheets[sheetName];
        const data = xlsx.utils.sheet_to_json(worksheet, { defval: '' });
        
        console.log('📈 Найдено строк:', data.length);

        res.writeHead(200, {
            'Content-Type': 'text/html; charset=utf-8',
            'Transfer-Encoding': 'chunked'
        });

        res.write(`
        <h1>Файл успешно загружен: ${req.file.filename}</h1>
        <h2>Создание PDF:</h2>
        <table id="pdf-table" border="1" cellspacing="0" cellpadding="5" style="border-collapse: collapse; width: 100%;">
          <thead>
            <tr style="background-color: #f2f2f2;">
              <th><input type="checkbox" id="select-all" /></th>
              <th>№</th>
              <th>Комната</th>
              <th>Имя</th>
              <th>Вода</th>
              <th>Свет</th>
              <th>Всего</th>
              <th>Счёт</th>
              <th>Скачать</th>
            </tr>
          </thead>
          <tbody></tbody>
        </table>

        <button id="send-selected" style="margin-top:20px; padding:10px 20px; background:#2196F3; color:white; border:none; border-radius:5px;">
          Отправить выбранные по почте
        </button>

        <button onclick="window.location.href='${ROUTE_PREFIX}/download-all'" 
        style="margin-top:20px; padding:10px 20px; background:#4CAF50; color:white; border:none; border-radius:5px;">
          Скачать все счета ZIP
        </button>

        <script>
        let counter = 0;

        function addPdfRow(room, name, water, electricity, total, status, pdfPath) {
          counter++;
          const tbody = document.querySelector('#pdf-table tbody');
          const row = document.createElement('tr');

          const checkboxCell = '<td><input type="checkbox" class="pdf-checkbox" data-pdf="' + pdfPath + '" /></td>';

          const statusCell = '<td style="background:' + (status === 'success' ? '#c6efce' : '#ffc7ce') +
                             '; text-align:center; font-weight:bold;">' +
                             (status === 'success' ? 'SUCCESS' : 'ERROR') + '</td>';

          const downloadCell = status === 'success'
              ? '<td><a href="' + pdfPath + '" target="_blank" style="display:inline-block; padding:5px 10px; background:#4CAF50; color:white; text-decoration:none; border-radius:5px;">Скачать</a></td>'
              : '<td>-</td>';

          row.innerHTML = checkboxCell +
                          '<td>' + counter + '</td>' +
                          '<td>' + room + '</td>' +
                          '<td>' + name + '</td>' +
                          '<td>' + water + '</td>' +
                          '<td>' + electricity + '</td>' +
                          '<td>' + total + '</td>' +
                          statusCell +
                          downloadCell;

          tbody.appendChild(row);
          window.scrollTo(0, document.body.scrollHeight);
        }

        document.getElementById('select-all').addEventListener('change', function() {
          const checked = this.checked;
          document.querySelectorAll('.pdf-checkbox').forEach(cb => cb.checked = checked);
        });

        document.getElementById('send-selected').addEventListener('click', async () => {
          const selected = Array.from(document.querySelectorAll('.pdf-checkbox:checked'))
                                .map(cb => cb.getAttribute('data-pdf'));
          if (selected.length === 0) {
            alert('Выберите хотя бы один счет для отправки.');
            return;
          }

          const email = prompt('Введите email для отправки:');
          if (!email) return;

          const response = await fetch('${ROUTE_PREFIX}/send-selected', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ email, pdfs: selected })
          });

          const result = await response.json();
          if(result.success) alert('Письмо отправлено успешно!');
          else alert('Ошибка при отправке письма.');
        });

        </script>
        `);

        // Получаем браузер
        console.log('🖥️ Получаем экземпляр браузера...');
        const browser = await getBrowser();
        console.log('✅ Браузер готов к работе');
        
        let successCount = 0;
        let errorCount = 0;

        for (let rowIndex = 2; rowIndex < data.length; rowIndex++) {
            const row = data[rowIndex];
            const name = row['Guest name'] || '';
            const room = row['Room no.'] || '';
            const water_start = row['Water Meter numbers'] || '';
            const water_end = row['__EMPTY_2'] || '';
            const water_consumption = row['Water consumption'] || '';
            const water_price = 89;
            const water_total = row['__EMPTY_3'] || '0';
            const electricity_start = row['Electricity Meter numbers'] || '';
            const electricity_end = row['__EMPTY_4'] || '';
            const electricity_consumption = row['Eletricity'] || '0';
            const electricity_price = 8;
            const electricity_total = row['__EMPTY_5'] || '0';
            const amount_total = row['Before amount'] || '0';
            const amount_before_vat = row['Before amount'] || '0';
            const vat = row['SVC'] || '0';
            const amount_total_net = row['Total amount'] || '0';
            const date_from = excelDateToDDMMYYYY(row['Period Check']) || '';
            const date_to = excelDateToDDMMYYYY(row['__EMPTY_1']) || '';
            const total_in_thai = toThaiBahtText(amount_total_net)
            const total_in_english = toWords(amount_total_net)

            function excelDateToDDMMYYYY(serial) {
              const excelEpoch = new Date(Date.UTC(1899, 11, 30));
              const days = Math.floor(serial);
              const milliseconds = days * 24 * 60 * 60 * 1000;
              const date = new Date(excelEpoch.getTime() + milliseconds);
            
              const dd = String(date.getUTCDate()).padStart(2, '0');
              const mm = String(date.getUTCMonth() + 1).padStart(2, '0');
              const yyyy = date.getUTCFullYear();
            
              return `${dd}/${mm}/${yyyy}`;
            }

            if (!name && !room) continue;

            try {
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
                                         .replace('{{date_from}}', date_from)
                                         .replace('{{date_to}}', date_to)
                                         .replace('{{total_in_thai}}', total_in_thai)
                                         .replace('{{total_in_english}}', total_in_english)
                                         .replace('{{qr_base64}}', qrDataUri)
                                         .replace('{{logo_base64}}', logoDataUri);

                const page = await browser.newPage();
                await page.setContent(invoiceHtml, { waitUntil: 'networkidle0', timeout: 30000 });
                const pdfFileName = `${name.replace(/\s+/g, '_')}_${room}_${Date.now()}.pdf`;
                const pdfPath = path.join(pdfFolder, pdfFileName);
                await page.pdf({ path: pdfPath, format: 'A4', printBackground: true, timeout: 30000 });
                await page.close();
                const pdfUrl = `${ROUTE_PREFIX}/pdf/${pdfFileName}`;

                res.write(`<script>addPdfRow("${room}", "${name}", "${water_total}", "${electricity_total}", "${amount_total}", "success", "${pdfUrl}");</script>`);
                successCount++;
            } catch (error) {
                errorCount++;
                res.write(`<script>addPdfRow("${room}", "${name}", "${water_total}", "${electricity_total}", "${amount_total}", "error", "");</script>`);
            }
        }

        res.write(`<h3>Генерация завершена! Успешно: ${successCount}, Ошибок: ${errorCount}</h3>`);
        res.end();

    } catch (error) {
        console.error('❌ Критическая ошибка:', error);
        res.status(500).send('Ошибка: ' + error.message);
    }
});

// Завершение приложения
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
