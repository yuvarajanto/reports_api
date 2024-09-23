const mongoose = require('mongoose');
const {nseDB} = require('../db');

const newquoteSchema = new mongoose.Schema({
    reqId:{
    type:Number,
        required:true
    }, 
    pointDetails:{
        type:Array,
    }, 
    bandwidthDetails:{
        type:Array
    }, 
    vas_link:{
        type:Array
    },
    vas_link1:{
        type:Array
    },
    vas_link2:{
        type:Array
    }, 
    customerName:{
        type:String
    }, 
    ebsAccountNo:{
        type:String
    },
    noOfLinks:{
        type:String
    }, 
    pageTracker:{
        type:String
    },
    quoteType:{
        type:String
    }, 
    isActive:{
        type:Boolean
    },
    status:{
        type:String
    }, 
    createdDate:{
        type:Date,
        default:Date.now()
    },
    
    bandwidthPriceValue:{
        type:Object
    },
    routerCount:{
        type:Number
    },
    isShippinGst:{
        type:Boolean
    },
    isShippinGstB:{
        type:Boolean
    },
    isShippinGstB1:{
        type:Boolean
    },
    poDetails:{
        type:Object
    },
    routerPosition:{
        type:Number
    }
});

const newquote = nseDB.model('newquote',newquoteSchema)
module.exports = newquote