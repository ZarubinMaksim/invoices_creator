const express = require("express");
const multer = require("multer");
const path = require("path");
const fs = require("fs");
const cors = require("cors");
const xlsx = require('xlsx');
const toThaiBahtText = require('thai-baht-text');
const { toWords } = require('number-to-words');
const app = express();
const PORT = 4000;
const puppeteer = require("puppeteer");


app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

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


// Проверка, что сервер жив
app.get("/", (req, res) => {
  res.send("✅ Сервер работает!");
});

let browserInstance = null;

// Функция для получения экземпляра браузера
const getBrowser = async () => {
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

// ----------------ЗАГРУЗКА ДОКУМЕНТА---------------------

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

app.post("/upload", upload.single("excel"), async(req, res) => {
  if (!req.file) return res.status(400).send("Файл не загружен");

  console.log('📖 Читаем Excel файл...', req.file.path);
  const workbook = xlsx.readFile(req.file.path);
  console.log('✅ Файл прочитан успешно');
  const sheetName = workbook.SheetNames[workbook.SheetNames.length - 4];
  console.log('📑 Выбран лист:', sheetName);
  const worksheet = workbook.Sheets[sheetName];
  const data = xlsx.utils.sheet_to_json(worksheet, { defval: '' });
  console.log('📈 Найдено строк:', data.length);

  let results = []; 
  
  console.log('🖥️ Получаем экземпляр браузера...');
  // const browser = await getBrowser();
  console.log('✅ Браузер готов к работе');
  
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

    console.log(`📊 Обрабатываем строку ${rowIndex}:`);

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




  try {


    let successCount = 0;
    let errorCount = 0;
    let invoiceCount = 0

  res.send({ message: "Файл успешно загружен!", filename: req.file.filename });

  } catch {

  }


});
//-------------------------------------------------------------


app.listen(PORT, '0.0.0.0', () => {
  console.log(`✅ Invoices server запущен на порту ${PORT}`);
  // console.log(`📋 Доступно по: http://38.244.150.204:${PORT}`);
});