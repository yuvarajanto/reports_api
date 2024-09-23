const mongoose=require('mongoose');
const {networkDB} = require('../db');

const quoteccmodelSchema = new mongoose.Schema({
    
    reqId:{
        type:Number,
        required:true
    },
    dcLocation:{
        type:String
    },
    connectionType:{
        type:String
    },

    companyDetails:{
        type:Object
    },
     
    aendRecords:{
        type:Object
    },
     
    zendRecords:{
        type:Object
    },
    otc:{
        type:Number,
    },
    arc:{
        type:Number,
    },
    currency:{
        type:String
    },
    ebsAccountNo:{
        type:String
    },
    companyId:{
        type:String
    },
    status:{
        type:String
    },
    loafile:{
        type:String
    },
    orderRefNo:{
        type:String
    },
    isActive:{
        type:Boolean
    },
    createdDate:{
        type:Date,
        default:Date.now()
    },
    orderedDate:{
        type:Date,
    },
     
    erpPayload:{
        type:Object
    },
     
    erpResponse:{
        type:Object
    }
})

const quotecc = networkDB.model('quotecc',quoteccmodelSchema);
module.exports = quotecc;