const express = require('express');
require('./src/db');
const bodyParser = require('body-parser')
const cors = require('cors');
const generateAndSendExcel = require('./src/router/emailconf');
const cron = require('node-cron');

const PORT =5100;

const app = express();
app.use(bodyParser.json());
app.use(cors());

cron.schedule('03 12 * * *', (res) => {
    console.log('Running scheduled task to send Excel report...');
    generateAndSendExcel();  
  });

app.listen(PORT,()=>{
    console.log(`Server Started at port ${PORT}`)
});