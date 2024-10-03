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
  if(excelBuffer.length === 0){
    const mailOptions = {
    from: 'onesify@sifycorp.com', 
    to: ['kiran.sudharsan@sifycorp.com','vaishnavi.srinivasan@sifycorp.com','yuvaraj.subramanian@sifycorp.com','dinesh.dhanapalan@sifycorp.com','sugavanesh.mayavel@sifycorp.com','murali.janakiraman@sifycorp.com'],  
    subject: 'Report - Order Summary',   
    html:`
    <p>As of yesterday, there were no orders placed in the current quarter.</p>
    <p>Best regards,</p><p>Team OneSify</p>
  <div style="margin-top: 20px;">`
  }
  try {
    await transporter.sendMail(mailOptions);
    console.log('Email sent successfully!');
  } catch (error) {
    console.error('Error sending email:', error);
  }
}else{
  const mailOptions = {
    from: 'onesify@sifycorp.com', 
    to: ['kiran.sudharsan@sifycorp.com','vaishnavi.srinivasan@sifycorp.com','yuvaraj.subramanian@sifycorp.com','dinesh.dhanapalan@sifycorp.com','sugavanesh.mayavel@sifycorp.com','murali.janakiraman@sifycorp.com'],  
    subject: 'Report - Order Summary',   
    html:`
    <p>Please find the attached Order Summary Report .</p>
    <p>Best regards,</br>Team OneSify</p>
  <div style="margin-top: 20px;">`,  
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