var express = require('express');
var path = require('path');
var favicon = require('serve-favicon');
var logger = require('morgan');
var cookieParser = require('cookie-parser');
var bodyParser = require('body-parser');

var app = express();

// view engine setup
// uncomment after placing your favicon in /public
//app.use(favicon(path.join(__dirname, 'public', 'favicon.ico')));
app.use(logger('dev'));
app.use(bodyParser.json({limit: '5mb'}));
app.use(bodyParser.urlencoded({ extended: false }));
app.use(cookieParser());
app.use(express.static(path.join(__dirname, 'public')));

var creator = require('./creator.js');

app.post('/api/gen', function(req, res) {
  var body = req.body;
  var accessCode = body.accessCode;
  var email = body.email;
  var data = body.data;

  creator.createTo(__dirname, data, accessCode);

  res.sendStatus(200);
});


app.listen(8091);

module.exports = app;
