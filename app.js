const express = require('express');
const bodyParser = require('body-parser');
const fs = require('fs');
const path = require('path');
const mongoose = require('mongoose');
const multer = require('multer');

const adminRoutes = require('./routes/admin');
const shopRoutes = require('./routes/shop');
const userRoutes = require('./routes/user');
const User = require('./models/user');

let MongoClient = require('mongodb').MongoClient;
let url = "mongodb://localhost:27017/onlineshopping";

const excelToJson = require('convert-excel-to-json');

const app = express();

app.set('view engine', 'ejs');
mongoose.set('useCreateIndex', true);

app.use(bodyParser.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, 'public')));

app.use((req, res, next) => {
    User.findById('603754a20065662bccb362d1')
        .then(userInDB => {
            req.user = userInDB;
            next();
        })
        .catch(err => console.log(err));
});

app.use('/admin', adminRoutes);
app.use(shopRoutes);
app.use(userRoutes);

global.__basedir = __dirname;
 
// -> Multer Upload Storage
const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        cb(null, __basedir + '/uploads/')
    },
    filename: (req, file, cb) => {
        cb(null, file.fieldname + "-" + Date.now() + "-" + file.originalname)
    }
});
 
const upload = multer({storage: storage});
 
// -> Express Upload RestAPIs
app.post('/api/uploadfile', upload.single("uploadfile"), (req, res) =>{
    importExcelData2MongoDB(__basedir + '/uploads/' + req.file.filename);
    res.json({
        'msg': 'File uploaded/import successfully!', 'file': req.file
    });
});
 
// -> Import Excel File to MongoDB database
function importExcelData2MongoDB(filePath){
    // -> Read Excel File to Json Data
    const excelData = excelToJson({
        sourceFile: filePath,
        sheets:[{
            // Excel Sheet Name
            name: 'Products',
 
            // Header Row -> be skipped and will not be present at our result object.
            header:{
               rows: 1
            },
			
            // Mapping columns to keys
            columnToKey: {
                A: 'title',
                B: 'imageURL',
                C: 'price',
                D: 'stock',
                E: 'description'
            }
        }]
    });
 
    // -> Log Excel Data to Console
    console.log(excelData);

    // Insert Json-Object to MongoDB
    MongoClient.connect('mongodb://localhost:27017/onlineshopping', { useNewUrlParser: true }, (err, db) => {
        if (err) throw err;
        let dbo = db.db("onlineshopping");
        dbo.collection("products").insertMany(excelData.Products, (err, res) => {
            if (err) throw err;
            console.log("Number of documents inserted: " + res.insertedCount);
            /**
                Number of documents inserted: 5
            */
            db.close();
        });
    });
			
    fs.unlinkSync(filePath);
}

app.use((req, res, next) => {
    res.status(404).send('Page Not Found');
});

// app.use((err, req, res, next) => {
//     res.status(500).send('Something Broke!');
// });

mongoose.connect('mongodb://localhost:27017/onlineshopping', { useNewUrlParser: true, useUnifiedTopology: true })
    .then(() => {
        app.listen(3000, () => {
            console.log('Server is running on 3000');
        });
    })
    .catch(err => console.log(err));