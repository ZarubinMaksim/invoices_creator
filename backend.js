const express = require("express");
const multer = require("multer");
const path = require("path");
const fs = require("fs");
const cors = require("cors");
const xlsx = require('xlsx');
const toThaiBahtText = require('thai-baht-text');
const { toWords } = require('number-to-words');
const archiver = require('archiver');
const app = express();
const PORT = 4000;
const puppeteer = require("puppeteer");
const nodemailer = require('nodemailer');
require('dotenv').config();
const { execSync } = require('child_process');
const logoDataUri = `data:image/png;base64,${fs.readFileSync(path.join(__dirname, 'img/logo.png')).toString('base64')}`;
const qrDataUri = `data:image/png;base64,${fs.readFileSync(path.join(__dirname, 'img/qr.png')).toString('base64')}`;
const invoiceTemplate = fs.readFileSync(path.join(__dirname, 'invoice_template.html'), 'utf-8');


app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

console.log('🚀 Инициализация сервера...');

// Убиваем все висящие процессы Chromium перед запуском
console.log('🔄 Убиваем все процессы Chromium/Chrome...');
const clearChromiumProcesses = () => {
  try {
    execSync('pkill -f chromium', { stdio: 'ignore' });
    execSync('pkill -f chrome', { stdio: 'ignore' });    
    console.log('✅ Все процессы Chromium/Chrome завершены');
} catch (error) {
    console.log('ℹ️ Не было процессов Chromium/Chrome для завершения');
}
}
clearChromiumProcesses()



// Папка для сохранённых PDF
const pdfFolder = path.join(__dirname, 'saved_pdf');
if (!fs.existsSync(pdfFolder)) {
    console.log('📁 Создаем папку для PDF:', pdfFolder);
    fs.mkdirSync(pdfFolder, { recursive: true });
} else {
    console.log('📁 Папка для PDF уже существует:', pdfFolder);
}

// Префикс маршрута
// const ROUTE_PREFIX = '/invoices';

// Делаем папку доступной по URL
app.use('/pdf', express.static(path.join(__dirname, 'saved_pdf')));


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

// Транспорт для отправки Gmail (нужен app password)
const transporter = nodemailer.createTransport({
  host: "vps.lagreenhotel.com",
  port: 465,
  secure: true, // 465 требует SSL
  auth: {
    user: process.env.MAIL_USER,
    pass: process.env.MAIL_PASS, // тот же пароль, что в Outlook
  },

  // service: 'gmail',
  // auth: {ввв
  //     user: process.env.GMAIL_USER,
  //     pass: process.env.GMAIL_PASS  // не обычный пароль, а пароль приложения Google
  // }
});

// API для отправки писем
app.post('/send-emails', express.json(), async (req, res) => {
  const rows = Array.isArray(req.body.rows) ? req.body.rows : [];
  const results = [];

  for (const row of rows) {
    try {
      if (!row.date) throw new Error('No date provided');
      if (!row.email) throw new Error('No email provided');
      if (!row.pdf) throw new Error('No PDF path provided');

      /**
       * row.date приходит в формате DD/MM/YYYY
       */
      const baseDate = new Date(row.date.split('/').reverse().join('-'));

      // --- месяц для subject (расчётный месяц)
      const monthNameSubject = baseDate.toLocaleString('en-US', { month: 'long' });
      const yearSubject = baseDate.getFullYear();

      // --- месяц для текста письма (следующий месяц)
      const dueDate = new Date(baseDate);
      dueDate.setMonth(dueDate.getMonth() + 1);
      const monthNameText = dueDate.toLocaleString('en-US', { month: 'long' });
      const yearText = dueDate.getFullYear();

      /**
       * row.pdf:
       * /pdf/2026-01/invoice_001.pdf
       */
       const cleanRelativePath = decodeURIComponent(
        row.pdf.replace(/^\/pdf\//, '')
      );
       const absolutePdfPath = path.join(__dirname, 'saved_pdf', cleanRelativePath);

      if (!fs.existsSync(absolutePdfPath)) {
        throw new Error(`PDF not found: ${absolutePdfPath}`);
      }

      await transporter.sendMail({
        from: '"La Green Hotel & Residence" <juristic@lagreenhotel.com>',
        to: row.email,
        bcc: 'juristic@lagreenhotel.com',
        subject: `${row.room} Utility Charges Invoice in ${monthNameSubject} ${yearSubject}`,
        html: `
          <p>Dear ${row.name},</p>

          <p>Good morning from Juristic Person Condominium,<br>
          I hope this message finds you well.</p>

          <p>
            We are writing to inform you that the invoice for the utility charges related to your condominium unit has been issued.
            The invoice includes a detailed breakdown of the charges for the specified billing period, and the payment due date is
            <strong>12th ${monthNameText} ${yearText}</strong>.
          </p>

          <p>
            Once you have made the payment, please send us the payment slip via email to:
            <a href="mailto:juristic@lagreenhotel.com">juristic@lagreenhotel.com</a>
            or via WhatsApp no. +66924633222
          </p>

          <p>
            Should you have any questions or require clarification regarding the invoice,
            please do not hesitate to contact us. We are here to assist you and ensure that
            all your inquiries are promptly addressed.
          </p>

          <p>Thank you for your attention to this matter. Have a good day.</p>

          <p>
            Best regards,<br>
            Sumolthip Kraisuwan<br>
            Assistant of Juristic Person Manager<br>
            <img src="cid:sign" alt="Signature" style="width:750px; height:200px;" />
          </p>
        `,
        attachments: [
          {
            filename: path.basename(absolutePdfPath),
            path: absolutePdfPath
          },
          {
            filename: 'sign.png',
            path: path.join(__dirname, 'img', 'sign.png'),
            cid: 'sign'
          }
        ]
      });

      results.push({ id: row.id, status: 'success' });

    } catch (err) {
      console.error('❌ Ошибка отправки на', row.email, err.message);
      results.push({ id: row.id, status: 'error', message: err.message });
    }
  }

  res.json({ results });
});



// Глобальная переменная для браузера
let browserInstance = null;

async function getBrowser() {
  if (!browserInstance) {
      console.log('🌐 Запускаем браузер...');
      
      const launchOptions = {
          args: [
              '--no-sandbox',
              '--disable-setuid-sandbox',
              '--disable-dev-shm-usage',
              '--disable-gpu',
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

function getMonthNameFromExcelDate(serial) {
  const excelEpoch = new Date(Date.UTC(1899, 11, 30)); // база для Excel
  const days = Math.floor(serial);
  const milliseconds = days * 24 * 60 * 60 * 1000;
  const date = new Date(excelEpoch.getTime() + milliseconds);
  
  // Получаем название месяца (например, "January", "February")
  return date.toLocaleString('en-US', { month: 'long' });
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


// для отправки логов в реальном времени
let clients = [];

app.get('/events', (req, res) => {
  res.setHeader('Content-Type', 'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache');
  res.setHeader('Connection', 'keep-alive');

  // добавляем клиента в массив
  clients.push(res);

  req.on('close', () => {
    clients = clients.filter(c => c !== res);
  });
});

function sendLog(message) {
  clients.forEach(res => {
    res.write(`data: ${JSON.stringify({ message })}\n\n`);
  });
}

async function closeBrowser() {
  if (browserInstance) {
    try { await browserInstance.close(); } catch (e) { console.error(e); }
    browserInstance = null;
  }
}


app.post(`/upload`, upload.single('excel'), async (req, res) => {
  console.log('📤 Получен POST запрос на загрузку файла');
  sendLog('📤 Uploading')

  if (!req.file) {
      console.log('❌ Файл не загружен');
      sendLog('❌ Error. File did not upload')
      return res.status(400).send('Файл не загружен');
  }

  console.log('✅ Файл загружен:', req.file.filename);
  sendLog('✅ File uploaded')


  try {
      console.log('📖 Читаем Excel файл...');
      sendLog('📖 Reading Excel file...')
      const workbook = xlsx.readFile(req.file.path);
      fs.unlink(req.file.path, () => {});
      console.log('✅ Файл прочитан успешно');
      sendLog('✅ Finish reading')
      
      const sheetIndex = workbook.SheetNames.length - 3;
      const sheetName = workbook.SheetNames[sheetIndex];

      // Берём последний лист (депозит)
      const depositIndex = workbook.SheetNames.length - 1;
      const depositName = workbook.SheetNames[depositIndex];


      console.log('📑 Выбран лист:', sheetName);
      sendLog('📑 Selected page:', sheetName)
      
      const worksheet = workbook.Sheets[sheetName];
      const depostSheet = workbook.Sheets[depositName];
      const data = xlsx.utils.sheet_to_json(worksheet, { defval: '' });
      
      // Очищаем папку saved_pdf (удаляем всё содержимое, саму папку оставляем)
      for (const file of fs.readdirSync(pdfFolder)) {
        fs.rmSync(path.join(pdfFolder, file), { recursive: true, force: true });
      }

      console.log('📁 Папка PDF очищена:', pdfFolder);
      sendLog('📁 PDF folder cleaned');

      const depositData = xlsx.utils.sheet_to_json(depostSheet, { defval: '' })
      console.log('📈 Найдено строк:', data.length);

    
      // Получаем браузер
      console.log('🖥️ Получаем экземпляр браузера...');
      sendLog('🔄 Starting PDF editor')
      const browser = await getBrowser();
      console.log('✅ Браузер готов к работе');
      
      let successCount = 0;
      let errorCount = 0;
      let invoiceCount = 0
      let results = []

      // создаём словарь депозитов
      // создаём словарь депозитов
      const depositMap = {};
depositData.forEach((row, index) => {
  if (index < 1) return;

  const rawRoom = row['Room no.'];
if (!rawRoom || typeof rawRoom !== 'string') return; // пропускаем, если нет строки

const roomNo = rawRoom
  .replace(/С/g, 'C') // русская С → английская C
  .replace(/В/g, 'B'); // русская В → английская B
  
  let deposit = parseFloat(row['__EMPTY_11']) || 0;

  depositMap[roomNo] = deposit;
});



      for (let rowIndex = 1; rowIndex < data.length; rowIndex++) { //it was rowIndex < data.length
          invoiceCount += 1
          const row = data[rowIndex];
          const name = row['Guest name'] || '';
          const room = row['Room no.'] || '';
          const deposit = (parseFloat(depositMap[room]) || 0).toFixed(2);
          const email = row['__EMPTY_1'] || '';
          const phone = row['__EMPTY_2'] || ''; //удалить когда колонки емаил и тел будут отдельные
          // const email = rawEmail.split(/[\s/]/)[0].trim();     //удалить когда колонки емаил и тел будут отдельные        
          // const email = '89940028777@ya.ru'
          const water_start = (parseFloat(row['Water Meter numbers']) || 0).toFixed(2);
          const water_end = (parseFloat(row['__EMPTY_4']) || 0).toFixed(2);
          const water_consumption = (parseFloat(row['Water consumption']) || 0).toFixed(2);
          const water_price = 89;
          const water_total = (parseFloat(row['__EMPTY_5']) || 0).toFixed(2);
          const electricity_start = (parseFloat(row['Electricity Meter numbers']) || 0).toFixed(2);
          const electricity_end = (parseFloat(row['__EMPTY_6']) || 0).toFixed(2);
          const electricity_consumption = (parseFloat(row['Eletricity']) || 0).toFixed(2);
          const electricity_price = 6.5;
          const electricity_total = (parseFloat(row['__EMPTY_7']) || 0).toFixed(2);
          const amount_total = (parseFloat(row['Before amount']) || 0).toFixed(2);
          const amount_before_vat = (parseFloat(row['Before amount']) || 0).toFixed(2);
          const vat = (parseFloat(row['SVC']) || 0).toFixed(2);
          const amount_total_net = (parseFloat(row['Total amount']) || 0).toFixed(2);
          const invoice_number = generateInvoiceNumber(invoiceCount, row['Period Check']); 
          const date_from = excelDateToDDMMYYYY(row['Period Check']) || '';
          const date_to = excelDateToDDMMYYYY(row['__EMPTY_3']) || '';
          const monthName = getMonthNameFromExcelDate(row['Period Check']);
          const isPaid = row['Paid'] || '';
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
          let page = null;

          try {
              console.log('📄 Читаем HTML шаблон...');
              sendLog('📄 Reading template...')
              // const logoPath = path.join(__dirname, 'img/logo.png');
              // const qrPath = path.join(__dirname, 'img/qr.png');
              // const logoBase64 = fs.readFileSync(logoPath).toString('base64');
              // const qrBase64 = fs.readFileSync(qrPath).toString('base64');
              let invoiceHtml = invoiceTemplate;
              // const logoDataUri = `data:image/png;base64,${logoBase64}`;
              // const qrDataUri = `data:image/png;base64,${qrBase64}`;
              // let invoiceHtml = fs.readFileSync(path.join(__dirname, 'invoice_template.html'), 'utf-8');
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
              sendLog('🆕 Creating new page...')
              page = await browser.newPage();              
              console.log('🔄 Устанавливаем контент...');
              sendLog('🔄 Setting up content...')
              await page.setContent(invoiceHtml, { 
                  waitUntil: 'load',
                  timeout: 30000
              });
              const pdfFileName = `${room}_${monthName}_${name.replace(/\s+/g, '_')}_${invoice_number}.pdf`;
              const pdfPath = path.join(pdfFolder, pdfFileName);
              console.log('🖨️ Генерируем PDF:', pdfPath);
              sendLog('🖨️ Creating PDF:', pdfPath)
              
              await page.pdf({ 
                  path: pdfPath, 
                  format: 'A4', 
                  printBackground: true,
                  timeout: 30000
              });
              
              console.log('✅ PDF успешно создан');
              sendLog(`✅ PDF has been created!: ${pdfFileName}`);
              const pdfUrl = `/pdf/${encodeURIComponent(pdfFileName)}`;
              results.push({
                room,
                name,
                email,
                phone,
                water_total,
                electricity_total,
                amount_total,
                status: 'success',
                deposit,
                isPaid,
                pdfUrl: pdfUrl,
                date_from
            });

              successCount++;
              
          } catch (error) {
            console.error('❌ Ошибка генерации PDF для строки', rowIndex, error);
            sendLog('❌ Error for row - ', rowIndex, error)
            results.push({
                room,
                name,
                email,
                phone,
                water_total,
                electricity_total,
                amount_total,
                status: 'error',
                deposit,
                isPaid,
                pdfUrl: null,
                date_from
            })
              errorCount++;
            } finally {
            if (page) await page.close().catch(() => {});
            }
      }
      res.json({ results });
      console.log(`✅ Обработка завершена. Успешно: ${successCount}, Ошибок: ${errorCount}`);
      sendLog(`✅ Finished. Successfull: ${successCount}, Errors: ${errorCount}`)
  } catch (error) {
      console.error('❌ Критическая ошибка:', error);
      sendLog('❌ Fatal error:', error)
      res.status(500).send('Ошибка: ' + error.message);
  } finally {
  await closeBrowser();
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

//all invoices ZIP
app.get('/download-all', (req, res) => {
  const zipName = `all_invoices_${Date.now()}.zip`;
  res.setHeader('Content-Disposition', `attachment; filename=${zipName}`);
  res.setHeader('Content-Type', 'application/zip');

  const archive = archiver('zip', { zlib: { level: 9 } });

  archive.on('error', err => {
    console.error('❌ Ошибка архивации:', err);
    res.status(500).send({ error: err.message });
  });

  archive.pipe(res);
  archive.directory(pdfFolder, false); // 🔥 ВАЖНО
  archive.finalize();
});


//download selected
app.post('/download-selected', express.json(), (req, res) => {
  const { pdfUrls } = req.body;
  console.log(pdfUrls)

  if (!Array.isArray(pdfUrls) || pdfUrls.length === 0) {
    return res.status(400).json({ error: 'Нет выбранных файлов' });
  }

  const zipName = `selected_invoices_${Date.now()}.zip`;

  res.setHeader('Content-Disposition', `attachment; filename="${zipName}"`);
  res.setHeader('Content-Type', 'application/zip');

  const archive = archiver('zip', { zlib: { level: 9 } });

  archive.on('error', (err) => {
    console.error('❌ Ошибка архивации:', err);
    if (!res.headersSent) {
      res.status(500).json({ error: 'Ошибка при создании архива' });
    }
    res.destroy();
  });

  // если клиент закрыл соединение — останавливаем архив
  // req.on('close', () => {
  //   archive.abort();
  // });

  archive.pipe(res);

  const basePdfFolder = path.join(process.cwd(), 'saved_pdf');

  pdfUrls.forEach((url) => {
    console.log('YAYAYAYYYA', url)
    /**
     * url приходит в виде:
     * /pdf/2026-01/invoice_001.pdf
     */
     const cleanRelativePath = decodeURIComponent(
      url.replace(/^\/pdf\//, '')
    );
    const absoluteFilePath = path.join(basePdfFolder, cleanRelativePath);
    const nameInZip = path.basename(cleanRelativePath);

    console.log('TRY:', absoluteFilePath);
    console.log('EXISTS:', fs.existsSync(absoluteFilePath));

    if (fs.existsSync(absoluteFilePath)) {
      archive.file(absoluteFilePath, { name: nameInZip });
    } else {
      console.warn('⚠️ Файл не найден:', absoluteFilePath);
    }
  });

  archive.finalize();
});

// Функция для генерации одного PDF на лету (без сохранения на сервере)
app.post('/generate-single-pdf', express.json(), async (req, res) => {
  console.log('📤 Получен запрос на генерацию одного PDF');
  let page = null
  try {
    const invoiceData = req.body;
    console.log(invoiceData)
    // Проверяем обязательные поля
    if (!invoiceData || !invoiceData.name || !invoiceData.room) {
      return res.status(400).json({ 
        error: 'Необходимы name и room для генерации PDF' 
      });
    }

    console.log('📊 Данные для генерации:', {
      name: invoiceData.name,
      room: invoiceData.room,
      date_from: invoiceData.date_from,
      date_to: invoiceData.date_to
    });

    // Подготавливаем данные
    const {
      name = '',
      room = '',
      water_start = '0',
      water_end = '0',
      water_consumption = '0',
      water_price = 89,
      water_total = '0',
      electricity_start = '0',
      electricity_end = '0',
      electricity_consumption = '0',
      electricity_price = 6.5,
      electricity_total = '0',
      amount_total = '0',
      amount_before_vat = '0',
      vat = '0',
      amount_total_net = '0',
      invoice_number = 'TEMP-INV-' + Date.now(),
      date_from = getCurrentDate(),
      date_to = getCurrentDate(),
      date_of_creating = getCurrentDate(),
    } = invoiceData;

    // Если не переданы суммы на тайском и английском, вычисляем их
    const finalAmountTotalNet = parseFloat(amount_total_net) || 0;
    const finalTotalInThai = toThaiBahtText(finalAmountTotalNet);
    const finalTotalInEnglish = toWords(finalAmountTotalNet);

    // Загружаем изображения
    // const logoPath = path.join(__dirname, 'img/logo.png');
    // const qrPath = path.join(__dirname, 'img/qr.png');
    // const logoBase64 = fs.readFileSync(logoPath).toString('base64');
    // const qrBase64 = fs.readFileSync(qrPath).toString('base64');
    // const logoDataUri = `data:image/png;base64,${logoBase64}`;
    // const qrDataUri = `data:image/png;base64,${qrBase64}`;
    let invoiceHtml = invoiceTemplate;
    // Читаем шаблон
    // let invoiceHtml = fs.readFileSync(
    //   path.join(__dirname, 'invoice_template.html'), 
    //   'utf-8'
    // );

    // Заменяем плейсхолдеры
    invoiceHtml = invoiceHtml
      .replace('{{name}}', name)
      .replace('{{room}}', room)
      .replace('{{water_start}}', parseFloat(water_start).toFixed(2))
      .replace('{{water_end}}', parseFloat(water_end).toFixed(2))
      .replace('{{water_consumption}}', parseFloat(water_consumption).toFixed(2))
      .replace('{{water_price}}', water_price)
      .replace('{{water_total}}', parseFloat(water_total).toFixed(2))
      .replace('{{electricity_start}}', parseFloat(electricity_start).toFixed(2))
      .replace('{{electricity_end}}', parseFloat(electricity_end).toFixed(2))
      .replace('{{electricity_consumption}}', parseFloat(electricity_consumption).toFixed(2))
      .replace('{{electricity_price}}', electricity_price)
      .replace('{{electricity_total}}', parseFloat(electricity_total).toFixed(2))
      .replace('{{amount_total}}', parseFloat(amount_total).toFixed(2))
      .replace('{{amount_before_vat}}', parseFloat(amount_before_vat).toFixed(2))
      .replace('{{vat}}', parseFloat(vat).toFixed(2))
      .replace('{{amount_total_net}}', parseFloat(amount_total_net).toFixed(2))
      .replace('{{invoice_number}}', invoice_number)
      .replace('{{date_from}}', date_from)
      .replace('{{date_to}}', date_to)
      .replace('{{date_of_creating}}', date_of_creating)
      .replace('{{total_in_thai}}', finalTotalInThai)
      .replace('{{total_in_english}}', finalTotalInEnglish)
      .replace('{{qr_base64}}', qrDataUri)
      .replace('{{logo_base64}}', logoDataUri);

    // Создаем PDF в памяти
    console.log('🖥️ Запускаем браузер для генерации PDF...');
    const browser = await getBrowser();
    page = await browser.newPage();
    
    await page.setContent(invoiceHtml, { 
      waitUntil: 'load',
      timeout: 30000
    });

    // Генерируем PDF в буфер
    console.log('🖨️ Генерируем PDF в буфер...');
    const pdfBuffer = await page.pdf({ 
      format: 'A4', 
      printBackground: true,
      timeout: 30000
    });

    await page.close();
    
    // Формируем имя файла
    const fileName = `${room}_${name.replace(/\s+/g, '_')}_${invoice_number}.pdf`
      .replace(/[^a-zA-Z0-9_.-]/g, '_');
    
    // Отправляем PDF прямо в браузер
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `inline; filename="${fileName}"`);
    res.setHeader('Content-Length', pdfBuffer.length);
    
    console.log(`✅ PDF успешно сгенерирован и отправлен: ${fileName} (${pdfBuffer.length} bytes)`);
    
    // Не закрываем браузер, чтобы он мог использоваться дальше
    // (браузер управляется через getBrowser и будет закрыт при SIGINT)
    
    res.end(pdfBuffer);
    
  } catch (error) {
    console.error('❌ Ошибка генерации PDF:', error);
    
    if (!res.headersSent) {
      res.status(500).json({ 
        error: 'Ошибка генерации PDF', 
        details: error.message 
      });
    }
  } finally {
    if (page) await page.close().catch(() => {});
}
});

app.listen(PORT, '0.0.0.0', () => {
  console.log(`✅ Invoices server запущен на порту ${PORT}`);
  // console.log(`📋 Доступно по: http://38.244.150.204:${PORT}`);
});

