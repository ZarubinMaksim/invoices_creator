const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const xlsx = require('xlsx');
const app = express();
const PORT = 4000;
const puppeteer = require('puppeteer');
const { execSync } = require('child_process');

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

// Маршрут для загрузки файла
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
            <ul id="pdf-list"></ul>
            <script>
                function addPdfItem(text) {
                    const list = document.getElementById('pdf-list');
                    const item = document.createElement('li');
                    item.textContent = text;
                    list.appendChild(item);
                    window.scrollTo(0, document.body.scrollHeight);
                }
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
            console.log('ROW ROW ROW', row)
            const name = row['Guest name'] || '';
            const room = row['Room no.'] || '';
            const water_start = row['Room no.'] || '';

            console.log(`📊 Обрабатываем строку ${rowIndex}:`, { name, room, amount });

            if (!name && !room && !amount) {
                console.log('⏭️ Пропускаем пустую строку');
                continue;
            }

            try {
                console.log('📄 Читаем HTML шаблон...');
                const logoPath = path.join(__dirname, 'img/logo.png');
                const logoBase64 = fs.readFileSync(logoPath).toString('base64');
                const logoDataUri = `data:image/png;base64,${logoBase64}`;
                let invoiceHtml = fs.readFileSync(path.join(__dirname, 'invoice_template.html'), 'utf-8');
                invoiceHtml = invoiceHtml.replace('{{name}}', name)
                                         .replace('{{room}}', room)
                                         .replace('{{amount}}', amount)
                                         .replace('{{logo_base64}}', logoDataUri);
;

                // Создаем новую страницу
                console.log('🆕 Создаем новую страницу...');
                const page = await browser.newPage();
                
                console.log('🔄 Устанавливаем контент...');
                await page.setContent(invoiceHtml, { 
                    waitUntil: 'networkidle0',
                    timeout: 30000
                });

                const pdfPath = path.join(pdfFolder, `${name.replace(/\s+/g, '_')}_${room}_${Date.now()}.pdf`);
                console.log('🖨️ Генерируем PDF:', pdfPath);
                
                await page.pdf({ 
                    path: pdfPath, 
                    format: 'A4', 
                    printBackground: true,
                    timeout: 30000
                });
                
                console.log('✅ PDF успешно создан');
                await page.close();
                
                res.write(`<script>addPdfItem("${name} - ${room} - ${pdfPath}");</script>`);
                successCount++;
                
            } catch (error) {
                console.error('❌ Ошибка:', error);
                errorCount++;
                res.write(`<script>addPdfItem("ОШИБКА: ${name} - ${room}");</script>`);
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