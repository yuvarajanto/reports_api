const mongoose = require('mongoose');

const nseDB = mongoose.createConnection(
    `mongodb+srv://kirankumar73056:kiran@cluster0.egvbtkw.mongodb.net/onesify_nse?retryWrites=true&w=majority&appName=Cluster0`,
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
    `mongodb+srv://kirankumar73056:kiran@cluster0.egvbtkw.mongodb.net/network_demo?retryWrites=true&w=majority&appName=Cluster0`,
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