const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = 4000;

// Префикс маршрута
const ROUTE_PREFIX = '/invoices';

// Папка для загрузок
const uploadFolder = path.join(__dirname, 'uploads');
if (!fs.existsSync(uploadFolder)) fs.mkdirSync(uploadFolder);

// Настройка multer
const storage = multer.diskStorage({
    destination: (req, file, cb) => cb(null, uploadFolder),
    filename: (req, file, cb) => {
        const ext = path.extname(file.originalname);
        const name = path.basename(file.originalname, ext);
        cb(null, `${name}-${Date.now()}${ext}`);
    }
});
const upload = multer({ storage });

// Страница с формой загрузки
app.get(`${ROUTE_PREFIX}`, (req, res) => {
    res.send(`
        <h2>Загрузка файла</h2>
        <form action="${ROUTE_PREFIX}/upload" method="post" enctype="multipart/form-data">
            <input type="file" name="document" required />
            <button type="submit">Загрузить</button>
        </form>
    `);
});

// Маршрут для загрузки файла
app.post(`${ROUTE_PREFIX}/upload`, upload.single('document'), (req, res) => {
    if (!req.file) return res.status(400).send('Файл не загружен');

    console.log('Файл загружен:', req.file.path);
    res.send(`Файл успешно загружен: ${req.file.filename}`);
});

// Слушаем все внешние подключения
app.listen(PORT, '0.0.0.0', () => {
    console.log(`Сервер запущен на http://38.244.150.204:${PORT}${ROUTE_PREFIX}`);
});
