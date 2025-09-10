const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const xlsx = require('xlsx');
const app = express();
const PORT = 4000;
const puppeteer = require('puppeteer');

console.log('🚀 Инициализация сервера...');

// Папка для сохранённых PDF
const pdfFolder = path.join(__dirname, 'saved_pdf');
if (!fs.existsSync(pdfFolder)) {
    console.log('📁 Создаем папку для PDF:', pdfFolder);
    fs.mkdirSync(pdfFolder);
} else {
    console.log('📁 Папка для PDF уже существует:', pdfFolder);
}

// Префикс маршрута
const ROUTE_PREFIX = '/invoices';

// Папка для загрузки файлов
const uploadFolder = path.join(__dirname, 'uploads');
if (!fs.existsSync(uploadFolder)) {
    console.log('📁 Создаем папку для загрузок:', uploadFolder);
    fs.mkdirSync(uploadFolder);
} else {
    console.log('📁 Папка для загрузок уже существует:', uploadFolder);
}

// Настройка multer с лимитом 100 MB
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
    limits: { fileSize: 100 * 1024 * 1024 } // 100 MB
});

// Глобальная переменная для браузера
let browserInstance = null;

// Функция для получения экземпляра браузера (синглтон)
async function getBrowser() {
    if (!browserInstance) {
        console.log('🌐 Запускаем браузер...');
        browserInstance = await puppeteer.launch({ 
            args: ['--no-sandbox', '--disable-setuid-sandbox'], 
            executablePath: '/usr/bin/chromium-browser', 
            headless: true 
        });
        console.log('✅ Браузер успешно запущен');
    } else {
        console.log('🔁 Используем существующий экземпляр браузера');
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
    console.log('📊 Путь к файлу:', req.file.path);

    try {
        console.log('📖 Читаем Excel файл...');
        const workbook = xlsx.readFile(req.file.path);
        console.log('✅ Файл прочитан успешно');
        
        const sheetIndex = workbook.SheetNames.length - 4; // предпоследний лист
        const sheetName = workbook.SheetNames[sheetIndex];
        console.log('📑 Выбран лист:', sheetName, '(индекс:', sheetIndex, ')');
        
        const worksheet = workbook.Sheets[sheetName];
        const data = xlsx.utils.sheet_to_json(worksheet, { defval: '' });
        
        console.log('📈 Найдено строк:', data.length);
        console.log('🔍 Пример данных первой строки:', JSON.stringify(data[0]));

        // Отправляем начальный ответ, чтобы избежать таймаута
        res.writeHead(200, {
            'Content-Type': 'text/html; charset=utf-8',
            'Transfer-Encoding': 'chunked'
        });
        
        console.log('📨 Отправляем начальный HTML ответ');
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

        // Получаем браузер один раз для всех PDF
        console.log('🖥️ Получаем экземпляр браузера...');
        const browser = await getBrowser();
        console.log('✅ Браузер готов к работе');
        
        let successCount = 0;
        let errorCount = 0;
        let skipCount = 0;

        for (let rowIndex = 2; rowIndex < data.length; rowIndex++) {
            const row = data[rowIndex];
            const name = row['Guest name'] || '';
            const room = row['Room no.'] || '';
            const amount = row['Total amount'] || '';

            console.log(`\n📊 Обрабатываем строку ${rowIndex}:`, { name, room, amount });

            // Пропускаем пустые строки
            if (!name && !room && !amount) {
                console.log('⏭️ Пропускаем пустую строку');
                skipCount++;
                continue;
            }

            try {
                console.log('📄 Читаем HTML шаблон...');
                let invoiceHtml = fs.readFileSync(path.join(__dirname, 'invoice_template.html'), 'utf-8');
                invoiceHtml = invoiceHtml.replace('{{name}}', name)
                                         .replace('{{room}}', room)
                                         .replace('{{amount}}', amount);
                console.log('✅ Шаблон подготовлен');

                // Создаем новую страницу для каждого PDF
                console.log('🆕 Создаем новую страницу в браузере...');
                const page = await browser.newPage();
                console.log('✅ Страница создана');

                console.log('🔄 Устанавливаем контент страницы...');
                await page.setContent(invoiceHtml, { waitUntil: 'networkidle0' });
                console.log('✅ Контент установлен');

                const pdfPath = path.join(pdfFolder, `${name.replace(/\s+/g, '_')}_${room}_${Date.now()}.pdf`);
                console.log('🖨️ Генерируем PDF:', pdfPath);
                
                await page.pdf({ 
                    path: pdfPath, 
                    format: 'A4', 
                    printBackground: true 
                });
                
                console.log('✅ PDF успешно создан');
                
                console.log('❌ Закрываем страницу...');
                await page.close();
                console.log('✅ Страница закрыта');
                
                // Отправляем информацию о созданном PDF клиенту
                console.log('📤 Отправляем информацию клиенту');
                res.write(`<script>addPdfItem("${name} - ${room} - ${pdfPath}");</script>`);
                console.log(`✅ Готово: ${room}, ${name}`);
                
                successCount++;
                
            } catch (error) {
                console.error(`❌ Ошибка при создании PDF для ${name} (${room}):`, error);
                res.write(`<script>addPdfItem("ОШИБКА: ${name} - ${room}");</script>`);
                errorCount++;
            }
        }

        console.log(`\n📊 ИТОГО: Успешно: ${successCount}, Ошибок: ${errorCount}, Пропущено: ${skipCount}`);

        // Завершаем ответ
        console.log('🏁 Завершаем ответ');
        res.write(`<h3>Генерация PDF завершена! Успешно: ${successCount}, Ошибок: ${errorCount}</h3>`);
        res.end();
        console.log('✅ Ответ завершен');

    } catch (error) {
        console.error('❌ Критическая ошибка при обработке файла:', error);
        if (!res.headersSent) {
            res.status(500).send('Ошибка при обработке файла: ' + error.message);
        } else {
            res.write(`<script>addPdfItem("Критическая ошибка: ${error.message}");</script>`);
            res.end();
        }
    } finally {
        console.log('🧹 Очистка завершена');
    }
});

// Обработка завершения приложения для закрытия браузера
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
    console.log(`✅ Invoices server запущен на http://38.244.150.204:${PORT}${ROUTE_PREFIX}`);
    console.log('📝 Логи процесса будут выводиться здесь');
    console.log('='.repeat(50));
});