const mongoose = require('mongoose');

const nseDB = mongoose.createConnection(
    'mongodb://127.0.0.1:27017/onesify_nse',
    {
        useNewUrlParser:true,
        useUnifiedTopology:true
    }
);


nseDB.on("error",()=>{
    console.error.bind(console," NSE_DB Connection Error");
});

nseDB.once("open",() => {
    console.info("NSE_DB Connected");
});

networkDB = mongoose.createConnection(
   'mongodb://127.0.0.1:27017/network_demo',
    {
        useNewUrlParser:true,
        useUnifiedTopology:true
    }
);

networkDB.on("error",()=>{
    console.error.bind(console," NETWORK_DB Connection Error");
});

networkDB.once("open",() => {
    console.info("NETWORJ_DB Connected");
});

module.exports = { nseDB, networkDB};