const express = require('express');
const cors = require('cors');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;

// Middleware
app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, 'frontend/dist/excel-processor-ui')));

// Serve static files from Angular build
app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'frontend/dist/excel-processor-ui/index.html'));
});

// Start server
app.listen(PORT, () => {
  console.log(`🚀 Excel Processor UI server running on port ${PORT}`);
  console.log(`📱 Frontend: http://localhost:${PORT}`);
});

module.exports = app;