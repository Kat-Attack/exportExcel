var express = require("express");
var excel = require("exceljs");
var fs = require('fs');
var app = express();

app.use(express.static(__dirname + "/public"));

app.get("/", function(req, res) {
    res.sendFile(__dirname + "/index.html");
});


app.post("/getexcel", function (req, res, next) {
    console.log("get excel file was requested.");
    var workbook = new excel.Workbook();
    var worksheet = workbook.addWorksheet("Sheet 1");

    worksheet.columns = [
        { header: "From", key: "from", width: 45 }, 
        { header: "To", key: "to", width: 45 }, 
        { header: "Block Number", key: "blocknumber", width: 12 }, 
        { header: "Hash", key: "hash", width: 70 }, 
        { header: "Value", key: "value", width: 8 }, 
        { header: "Input", key: "input",width: 40 }
    ];
        
    worksheet.addRow({
        from: "kathy",
        to: "you",
        blocknumber: "300",
        hash: "0x123ffgv4",
        value: "1000",
        input: "xbt",
    });
        
    workbook.xlsx.writeFile("./public/sample.xlsx").then(function (err) {
        console.log("xlsx file is written.");
        if (err) {
            console.log("error " + err);
            res.end();
        }

        //check if file is written first, and then send msg to client that it's ok to download'
        fs.exists("./public/sample.xlsx", function(exists) {
            if (exists) {
                console.log("file exists");
                res.end();
            
            }else{
                console.log("file doesn't exist");
            }
        });
    });            
});


app.listen(3000, function (err) {
    if (!err){
        console.log("Listening on 3000");
    } else {
        console.log("ERROR: " + err);
    }
})