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
    to: ['kiran.sudharsan@sifycorp.com','vaishnavi.srinivasan@sifycorp.com','yuvaraj.subramanian@sifycorp.com','dinesh.dhanapalan@sifycorp.com','sugavanesh.mayavel@sifycorp.com','murali.janakiraman@sifycorp.com','gomathi.sitaram@sifycorp.com','arun.rajamani@sifycorp.com'],
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
  const currDate = new Date();  
  const month = currDate.getMonth()+1;
  const queryQuarter = getQuarter(month);  
  console.log(formatDateAsDDMMYY(currDate));  
  const today = formatDateAsDDMMYY(currDate);
  const mailOptions = {
    from: 'onesify@sifycorp.com', 
    to: ['kiran.sudharsan@sifycorp.com','vaishnavi.srinivasan@sifycorp.com','yuvaraj.subramanian@sifycorp.com','dinesh.dhanapalan@sifycorp.com','sugavanesh.mayavel@sifycorp.com','murali.janakiraman@sifycorp.com','gomathi.sitaram@sifycorp.com','arun.rajamani@sifycorp.com'],  
    subject: 'Report - Order Summary',   
    html:`
    <p>Please find the attached Order Summary Report .</p>
    <p>Best regards,</br>Team OneSify</p>
  <div style="margin-top: 20px;">`,  
    attachments: [
      {
        filename: `OSP-OB-${queryQuarter}-${today}.xlsx`,  
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

  const formatDateAsDDMMYY = (date) => {
    const day = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const year = String(date.getFullYear()).slice(-2);
    return `${day}${month}${year}`; 
  };
 
  const getQuarter = (month) => { if (month >= 1 && month <= 3) return 'Q4'; if (month >= 4 && month <= 6) return 'Q1'; if (month >= 7 && month <= 9) return 'Q2'; return 'Q3';};

  module.exports = generateAndSendExcel;