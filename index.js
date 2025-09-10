const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const xlsx = require('xlsx');
const app = express();
const PORT = 4000;

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
app.post(`${ROUTE_PREFIX}/upload`, upload.single('excel'), (req, res) => {
    if (!req.file) return res.status(400).send('Файл не загружен');

    console.log('Файл загружен:', req.file.path);
    res.send(`Файл успешно загружен: ${req.file.filename}`);

    const workbook = xlsx.readFile(req.file.path);
    const sheetIndex = workbook.SheetNames.length - 3;
    const sheetName = workbook.SheetNames[sheetIndex];
    const worksheet = workbook.Sheets[sheetName];

    const data = xlsx.utils.sheet_to_json(worksheet, { defval: '' });

    console.log('EXCL', data)
});

// Слушаем все внешние подключения
app.listen(PORT, '0.0.0.0', () => {
    console.log(`Invoices server запущен на http://38.244.150.204:${PORT}${ROUTE_PREFIX}`);
});
