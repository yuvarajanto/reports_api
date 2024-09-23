const nodemailer = require('nodemailer');
const {generateExcel} = require('./excelconf');
require('dotenv').config();

const sendEmailWithAttachment = async (excelBuffer) => {
  const transporter = nodemailer.createTransport({
    host: process.env.SMTP_Mail_Host,
    port: process.env.SMTP_Mail_port,
    auth: { user: process.env.SMTP_TO_EMAIL, pass: process.env.SMTP_TO_PASSWORD },
    secureConnection: false,
    tls: { ciphers: 'SSLv3' }
});
  
    const mailOptions = {
      from: 'onesify@sifycorp.com', 
      to: ['kiran.sudharsan@sifycorp.com','vaishnavi.srinivasan@sifycorp.com','yuvaraj.subramanian@sifycorp.com'],  
      subject: 'Report - Order Summary',  
      text: 'Please find the attached Excel report.',  
      attachments: [
        {
          filename: 'report.xlsx',  
          content: excelBuffer,  
          contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'  
        }
      ]
    };
  
    try {
      await transporter.sendMail(mailOptions);
      console.log('Email sent successfully!');
    } catch (error) {
      console.error('Error sending email:', error);
    }
  };

  const generateAndSendExcel = async () => {
    try {
   
      const excelBuffer = await generateExcel();

      await sendEmailWithAttachment(excelBuffer);
    } catch (error) {
      console.error('Error generating or sending Excel file:', error);
    }
  };

  module.exports = generateAndSendExcel;