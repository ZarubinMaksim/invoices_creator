const express = require('express');

const app = express();
const PORT = 4000;

// Главная страница
app.get('/', (req, res) => {
    res.send(`
        <h1>Привет! Это проект Invoices Creator</h1>
        <p>Сервер работает и доступен через Nginx на /invoices/</p>
    `);
});

// Слушаем все внешние подключения
app.listen(PORT, '0.0.0.0', () => {
    console.log(`Invoices server запущен на http://38.244.150.204:${PORT}`);
});