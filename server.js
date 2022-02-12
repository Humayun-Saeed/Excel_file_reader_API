var express=require('express')
var nodemailer=require('nodemailer')
var excel=require('exceljs')
var app=express()
require('dotenv').config()
const send= async(Email)=>{
    
const transporter=nodemailer.createTransport({

    service:'gmail',
    auth:{

        user:'humayunsaeed267@gmail.com',
        pass:process.env.password
    }
})
var mailoptions={

    form:'humayunsaeed267@gmail.com',
    to:Email,
    subject:'test',
    Text:'hello we build a api which chose the emails form the excel file and send the emails: '
}
transporter.sendMail(mailoptions,(err,info)=>{

    if(err){
        console.log(err);
    }
    else{
        console.log(info.response);
    }
})

        }

app.get('/email',async(req,res)=>{
const read=new excel.Workbook()
read.xlsx.readFile('./Book1.xlsx').then(async()=>{

    const sheet=read.getWorksheet("Sheet1")
    for(var i=1;i<=sheet.rowCount;i++){

        console.log( sheet.getRow(i).getCell(1).value.text);
        console.log("------")
         send(sheet.getRow(i).getCell(1).value.text)
    }
})
})

app.listen(30000,()=>{
    console.log("server is running at this port: ");
})