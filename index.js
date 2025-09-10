const express = require('express');
const app = express();
const PORT = 4000;

// простой маршрут
app.get('/', (req, res) => {
  res.send('Hello from Express!');
});

app.listen(PORT, () => {
  console.log(`✅ Server running at http://localhost:${PORT}`);
});
