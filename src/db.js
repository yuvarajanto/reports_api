const mongoose = require('mongoose');

const nseDB =mongoose.createConnection('mongodb://223.30.223.188:27017/onesify_nse', 
  {
    auth: {
        username:'appusr',
        password:'AppPass123'
      },
    authSource:"admin",
    useUnifiedTopology: true,
    useNewUrlParser: true
  });
   


nseDB.on("error",()=>{
    console.error.bind(console," NSE_DB Connection Error");
});

nseDB.once("open",() => {
    console.info("NSE_DB Connected");
});

networkDB = mongoose.createConnection('mongodb://223.30.223.188:27017/network', 
  {
    auth: {
        username:'appusr',
        password:'AppPass123'
      },
    authSource:"admin",
    useUnifiedTopology: true,
    useNewUrlParser: true
  });
   

networkDB.on("error",()=>{
    console.error.bind(console," NETWORK_DB Connection Error");
});

networkDB.once("open",() => {
    console.info("NETWORK_DB Connected");
});

module.exports = { nseDB, networkDB};