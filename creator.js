var xl = require('excel4node');
var fs = require('fs');
var nodemailer = require('nodemailer');

var transporter = nodemailer.createTransport({
    host: 'smtp.mail.ru',
    port: 465,
    secure: true, // use SSL
    auth: {
        user: 'testLacostApp@mail.ru',
        pass: '40049513579A'
    }
});

module.exports = {
    createTo: function(__dirname, data, accessCode, email) {
        fs.writeFile('./data/' + accessCode + '/attempt_at_' +Date.now()+'.txt', data ,function(err){
            if(err){
                console.log("error is: " + err);
            }
        });

        var ItemImageColSize = 5;
        var ItemImageRowSize = 12;
        var MaxPerRow = 10;

        var genders = [];
        data.map(function(item) {return item.gender;})
            .forEach(function(gender) {
                if (genders.indexOf(gender) < 0) {
                    genders.push(gender);
                }
            });

        var codes = [];
        data.map(function(item) {return item.code;})
            .forEach(function(code) {
                if (codes.indexOf(code) < 0) {
                    codes.push(code);
                }
            });

        var sheets = {};
        var info = {};

        var wb = new xl.Workbook();

        genders.forEach(function(gender) {
            sheets[gender] = wb.addWorksheet(gender);
            info[gender] = {
                lastRowItemsCount: 0,
                lastRowCurrentCol: 1,
                currentRow: 1
            };
        });

        codes.forEach(function(code) {
            var colors = data.filter(function(item) {
                return item.code == code;
            });

            var fItem = colors[0];
            var fRow =info[fItem.gender].currentRow;

            if (info[fItem.gender].lastRowCurrentCol == 1) {
                sheets[fItem.gender].cell(fRow + ItemImageRowSize, 1).string('Name');
                sheets[fItem.gender].cell(fRow + ItemImageRowSize + 1, 1).string('Code');
                sheets[fItem.gender].cell(fRow + ItemImageRowSize + 2, 1).string('Color Name');
                sheets[fItem.gender].cell(fRow + ItemImageRowSize + 3, 1).string('Color Code');
                sheets[fItem.gender].cell(fRow + ItemImageRowSize + 4, 1).string('Code+Color');
                sheets[fItem.gender].cell(fRow + ItemImageRowSize + 5, 1).string('Material');
                sheets[fItem.gender].cell(fRow + ItemImageRowSize + 6, 1).string('Order');
                info[fItem.gender].lastRowCurrentCol++;
            }

            var firstCol = info[fItem.gender].lastRowCurrentCol;
            var lastCol = firstCol;

            colors.forEach(function(item) {
                if (!item.quantity || isNaN(item.quantity)) {
                    return;
                }

                var currentRow = info[item.gender].currentRow;
                var currentCol = info[item.gender].lastRowCurrentCol;

                //console.log(item.code);
                //console.log(item.gender);
                //console.log(currentRow + " " + currentCol);
                //console.log((currentRow + ItemImageRowSize) + " " + (currentCol + ItemImageColSize));

                sheets[item.gender].addImage({
                    path: __dirname + '/data/' + accessCode + '/' +item.url,
                    type: 'picture',
                    position: {
                        type: 'twoCellAnchor',
                        from: {
                            col: currentCol,
                            row: currentRow,
                            colOff: '0.5in'
                        },
                        to: {
                            col: currentCol + ItemImageColSize - 1,
                            row: currentRow + ItemImageRowSize - 2,
                            colOff: '0.5in'
                        }
                    }
                });
                sheets[item.gender]
                    .cell(
                        currentRow + ItemImageRowSize + 2, currentCol,
                        currentRow + ItemImageRowSize + 2, currentCol+ ItemImageColSize - 1,
                        true
                    )
                    .style({
                        alignment: {
                            wrapText: true,
                            horizontal: 'center'
                        }
                    })
                    .string(item.colorCode);
                sheets[item.gender]
                    .cell(
                        currentRow + ItemImageRowSize + 3, currentCol,
                        currentRow + ItemImageRowSize + 3, currentCol+ ItemImageColSize - 1,
                        true
                    )
                    .style({
                        alignment: {
                            wrapText: true,
                            horizontal: 'center'
                        }
                    })
                    .string(item.colorDesc);
                sheets[item.gender]
                    .cell(
                        currentRow + ItemImageRowSize + 4, currentCol,
                        currentRow + ItemImageRowSize + 4, currentCol+ ItemImageColSize - 1,
                        true
                    )
                    .style({
                        alignment: {
                            wrapText: true,
                            horizontal: 'center'
                        }
                    })
                    .string(item.code + item.colorCode);
                sheets[item.gender]
                    .cell(
                        currentRow + ItemImageRowSize + 5, currentCol,
                        currentRow + ItemImageRowSize + 5, currentCol+ ItemImageColSize - 1,
                        true
                    )
                    .style({
                        alignment: {
                            wrapText: true,
                            horizontal: 'center'
                        }
                    })
                    .string(item.material);

                sheets[item.gender]
                    .cell(
                        currentRow + ItemImageRowSize + 6, currentCol,
                        currentRow + ItemImageRowSize + 6, currentCol+ ItemImageColSize - 1,
                        true
                    )
                    .style({
                        alignment: {
                            wrapText: true,
                            horizontal: 'center'
                        }
                    })
                    .number(item.quantity);

                lastCol = currentCol+ ItemImageColSize - 1;
                info[item.gender].lastRowItemsCount++;
                info[item.gender].lastRowCurrentCol += ItemImageColSize;
            });

            sheets[fItem.gender]
                .cell(
                    fRow + ItemImageRowSize, firstCol,
                    fRow + ItemImageRowSize, lastCol,
                    true
                )
                .style({
                    alignment: {
                        wrapText: true,
                        horizontal: 'center'
                    }
                })
                .string(fItem.code);

            sheets[fItem.gender]
                .cell(
                    fRow + ItemImageRowSize + 1, firstCol,
                    fRow + ItemImageRowSize + 1, lastCol,
                    true
                )
                .style({
                    alignment: {
                        wrapText: true,
                        horizontal: 'center'
                    }
                })
                .string(fItem.name);

            //info[fItem.gender].currentCol ++;
            if (info[fItem.gender].lastRowItemsCount > MaxPerRow) {
                info[fItem.gender].lastRowCurrentCol = 1;
                info[fItem.gender].lastRowItemsCount = 0;
                info[fItem.gender].currentRow += ItemImageRowSize + 11;
            }
        });

        /*
        sheets['Man'].addImage({
            path: __dirname + '/placeholder.png',
            type: 'picture',
            position: {
                type: 'twoCellAnchor',
                from: {
                    col: 2,
                    row: 2
                },
                to: {
                    col: 6,
                    row: 10
                }
            }
        });
        */

        wb.write('./data/' + accessCode + '/ExcelFile.xlsx');

        setTimeout(
            function() {
                fs.readFile('./data/' + accessCode + '/ExcelFile.xlsx', function (err, data) {
                    transporter.sendMail({
                        sender: 'testLacostApp@mail.ru',
                        to: email,
                        subject: 'Attachment!',
                        body: 'Selected data here.',
                        attachments: [{'filename': 'data.xlsx', 'content': data}]
                    }, function (err, success) {
                        if (err) {
                            console.log(err);
                        }

                    });
                });
            }
            ,5000);
    }
};