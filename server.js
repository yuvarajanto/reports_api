const express = require('express');
require('./src/db');
require('dotenv').config();
const bodyParser = require('body-parser')
const cors = require('cors');
const generateAndSendExcel = require('./src/router/emailconf');
const cron = require('node-cron');
const https = require("https");



const app = express();
app.use(bodyParser.json());
app.use(cors());

cron.schedule('* 11 * * *', () => {
    console.log('Running scheduled task to send Excel report...');
    generateAndSendExcel();
  });

// HTTPS
var port = process.env.PORTWSSL;
var server = https.createServer(options, app);
server.listen(port, function () {
  console.log(
    `REPORTS API is running on port: ${process.env.PORTWSSL} at ${appDate}`
  );
});