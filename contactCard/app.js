var createError = require('http-errors');
var express = require('express');
var path = require('path');
var cookieParser = require('cookie-parser');
var logger = require('morgan');
const fileUpload = require('express-fileupload');
const vCardsJS = require('vcards-js');

var parser = require('simple-excel-to-json')

var indexRouter = require('./routes/index');
var usersRouter = require('./routes/users');

var app = express();

// view engine setup
app.set('views', path.join(__dirname, 'views'));
app.set('view engine', 'hbs');

app.use(logger('dev'));
app.use(express.json());
app.use(express.urlencoded({ extended: false }));
app.use(cookieParser());
app.use(express.static(path.join(__dirname, 'public')));



app.use('/', indexRouter);
app.use('/users', usersRouter);
app.use(fileUpload());

app.post('/upload', function (req, res) {
  if (!req.files || Object.keys(req.files).length === 0) {
    return res.status(400).send('No files were uploaded.');
  }

  // The name of the input field (i.e. "sampleFile") is used to retrieve the uploaded file
  let sampleFile = req.files.sampleFile;
  let fileName = sampleFile.name;
  let filePath = path.join(__dirname, 'tmp',sampleFile.name); 
  // Use the mv() method to place the file somewhere on your server
  sampleFile.mv(filePath, function(err) {
    if (err)
      return res.status(500).send(err);

    var contact = parser.parseXls2Json(filePath);  
    var response='';
    contact.forEach((sheet)=>{
      sheet.forEach((contact)=>{
        vCard = vCardsJS();
        vCard.firstName = contact.FirstName;
        vCard.lastName = contact.LastName;
        vCard.organization = contact.Organization;
        vCard.workPhone = contact.Number;
        response +=vCard.getFormattedString();
      })
    })

    res.set('Content-Type', 'text/vcard; name="contacts.vcf"');
    res.set('Content-Disposition', 'inline; filename="contacts.vcf"');
    res.send(response);
    
  
  });

 
});

// catch 404 and forward to error handler
app.use(function (req, res, next) {
  next(createError(404));
});

// error handler
app.use(function (err, req, res, next) {
  // set locals, only providing error in development
  res.locals.message = err.message;
  res.locals.error = req.app.get('env') === 'development' ? err : {};

  // render the error page
  res.status(err.status || 500);
  res.render('error');
});



module.exports = app;
