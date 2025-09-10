app.post(`${ROUTE_PREFIX}/upload`, upload.single('excel'), (req, res) => {
  if (!req.file) return res.status(400).send('Файл не загружен');

  const workbook = xlsx.readFile(req.file.path);
  const sheetIndex = workbook.SheetNames.length - 3; // предпоследний лист
  const sheetName = workbook.SheetNames[sheetIndex];
  const worksheet = workbook.Sheets[sheetName];

  const data = xlsx.utils.sheet_to_json(worksheet, { defval: '' });

  // Собираем HTML для вывода
  let html = `<h1>Файл успешно загружен: ${req.file.filename}</h1>`;
  html += `<h2>Данные из файла:</h2><ul>`;

  data.forEach((row, rowIndex) => {
      if (rowIndex < 2) return; // пропустить первые 2 строки
          const name = row['Guest name'] || 'N/A';
          const room = row['Room no.'] || 'N/A';
          const amount = row['Total amount'] || 'N/A';
          html += `<li>Owner data: Name - ${name}, Room - ${room}, Amount - ${amount}</li>`;
  });

  html += `</ul>`;
  res.send(html);
});
