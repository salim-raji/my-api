const express = require("express");
const xlsx = require("xlsx");
const app = express();
const fs = require('fs');

const nodemailer = require('nodemailer')

let transporter = nodemailer.createTransport({
    host: 'smtp.mail.me.com',
    port:587,
    secure: false,
    auth:{
        user: 'salim83102@icloud.com',
        pass: 'jjvl-wcyy-dcmm-klxq'
    },
});

let mailOptions = {
    from: 'salim83102@icloud.com',
    to: 'salimraji@icloud.com',
    subject: 'Testing',
    text:'Testing'
};

app.use(express.json());


//Midleware
const logRequest = (req, res, next) => {
    const timestamp = new Date().toString();

    const originalSend = res.send;

    const logEntry = {
        endpoint: req.url,
        method: req.method,
        time: timestamp,
        body: req.body
    };

    if(req.method === 'DELETE'){
        console.log('it is working')
    }

    if(req.method === 'DELETE'){
            
        transporter.sendMail(mailOptions, (error, info) => {
            if (error){
                return console.log(error);
            }else{
                console.log('email is ', info.response)
            }
        })
    }

    res.send = function (body) {
        logEntry.response = body;

        fs.appendFile('middlewareLogs.txt', JSON.stringify(logEntry) + '\n', (err) =>{
            if(err){
                console.error('Error writing log', err);
            }
        });
        console.log(req.method);
        return originalSend.call(this, body);
    };

    
    next();
}

app.use(logRequest)

const fileName = './myData.xlsx'
//API
if(!fs.existsSync(fileName)){
    const wb = xlsx.utils.book_new();
    const ws = xlsx.utils.json_to_sheet([]);
    xlsx.utils.book_append_sheet(wb, ws, 'Sheet1');
    xlsx.writeFile(wb,fileName);
}

const loadExcelFile = () => {
    const file = xlsx.readFile(fileName);
    var sheet_name_list = file.SheetNames;
    var xlData = xlsx.utils.sheet_to_json(file.Sheets[sheet_name_list[0]]);
    return xlData;
}

const saveExcelFile = (data) => {

    const ws = xlsx.utils.json_to_sheet(data);
    const wb = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(wb, ws, 'Sheet1');
    xlsx.writeFile(wb, fileName);
}


app.post ('/create' , (req,res) =>{
    const newData = req.body;
    const data = loadExcelFile();
    data.push(newData);
    saveExcelFile(data);
    res.json({message: 'Data added'})

});

app.get ('/read', (req,res) => {
    const data = loadExcelFile();
    res.json(data);
});

app.get ('/read/:id', (req,res) =>{
    const id = parseInt(req.params.id);
    let data = loadExcelFile();
    data = data.filter(item => item.id == id);
    res.json(data)
})

app.delete('/delete/:id', (req,res) =>{
    const id = parseInt(req.params.id);
    let data = loadExcelFile();
    data = data.filter(item => item.id !== id);
    saveExcelFile(data);
    res.json({message : 'Data Deleted'});
});

app.put('/update/:id', (req,res) => {
    const id = parseInt(req.params.id);
    const { firstName, lastName, email} = req.body;
    let data = loadExcelFile();
    const index = data.findIndex(item => item.id === id);
    if(firstName) data[index].firstName = firstName;
    if(lastName) data[index].lastName = lastName;
    if(email) data[index] = email;
    saveExcelFile(data);
    res.json({message: 'Data Updated'})

});

app.listen(3000, () => {
    console.log('Server running on port 3000')
});