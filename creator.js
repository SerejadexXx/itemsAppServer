var xl = require('excel4node');
var fs = require('fs');
var nodemailer = require('nodemailer');
var imageService = require('images');
var pathService = require('path');

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
    createTo: function(__dirname, data, comments, accessCode, email, name) {
            fs.writeFile('./data/' + accessCode + '/attempt_at_' + Date.now() + '.txt', JSON.stringify(data), function (err) {
                if (err) {
                    console.log("error is: " + err);
                }
            });
            fs.writeFile('./data/' + accessCode + '/attempt_com_at_' + Date.now() + '.txt', JSON.stringify(comments), function (err) {
                if (err) {
                    console.log("error is: " + err);
                }
            });

        var ItemImageColSize = 5;
        var ItemImageRowSize = 12;
        var MaxPerRow = 10;

        var pageNums = [];
        var GetPageNum = function(item) {
            var val = "";
            if (item.typeName) {
                val += item.typeName;
            }
            if (item.gender) {
                if (val == "") {
                    val += item.gender;
                } else {
                    val += " " + item.gender;
                }
            }
            return val;
        };
        data.forEach(function(item) {
                var val = GetPageNum(item);
                if (pageNums.indexOf(val) < 0) {
                    pageNums.push(val);
                }
            });

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

        var convertImage = function(path) {
            var saveUrl = pathService.dirname(path) + '/squared_' + pathService.basename(path);
            /*try {
                fs.accessSync(saveUrl, fs.F_OK);
            } catch (e) {*/
                var image = imageService(path);
                var width = image.width();
                var height = image.height();
                var size = Math.max(width, height);
                var rezImage = imageService(size, size);
                rezImage.fill(0xff, 0xff, 0xff, 1);
                rezImage.draw(image, Math.floor(size / 2 - width / 2), Math.floor(size / 2 - height / 2));
                rezImage.resize(200, 200);
                rezImage.save(
                    saveUrl,
                    {
                        quality: 100
                    }
                );
            //}
            return saveUrl;
        };

        var GetFormattedComments = function(code) {
            if (!comments) {
                return "";
            }
            var itemComments = comments.filter(function(one) {
                return one.code == code;
            });
            if (itemComments.length == 0) {
                return "";
            } else {
                var texts = itemComments[0].comments;
                var rez = "";
                texts.forEach(function(one) {
                    if (rez == "") {
                        rez = one.text;
                    } else {
                        rez += '\n\n' + one.text;
                    }
                });
                return rez;
            }
        };

        var wb = new xl.Workbook();

        pageNums.forEach(function(pageNum) {
            sheets[pageNum] = wb.addWorksheet(pageNum);
            info[pageNum] = {
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
            var pageNum = GetPageNum(fItem);
            var fRow =info[pageNum].currentRow;

            if (info[pageNum].lastRowCurrentCol == 1) {
                sheets[pageNum].cell(fRow + ItemImageRowSize, 1).string('Name');
                sheets[pageNum].cell(fRow + ItemImageRowSize + 1, 1).string('Code');
                sheets[pageNum].cell(fRow + ItemImageRowSize + 2, 1).string('Color Code');
                sheets[pageNum].cell(fRow + ItemImageRowSize + 3, 1).string('Color Name');
                sheets[pageNum].cell(fRow + ItemImageRowSize + 4, 1).string('Code+Color Code');
                sheets[pageNum].cell(fRow + ItemImageRowSize + 5, 1).string('Material');
                sheets[pageNum].cell(fRow + ItemImageRowSize + 6, 1).string('Order');
                sheets[pageNum].cell(fRow + ItemImageRowSize + 7, 1).string('Price');
                sheets[pageNum].cell(fRow + ItemImageRowSize + 8, 1).string('Comments');
                info[pageNum].lastRowCurrentCol++;
            }

            var firstCol = info[pageNum].lastRowCurrentCol;
            var lastCol = firstCol;

            colors.forEach(function(item) {
                if (!item.quantity || isNaN(item.quantity)) {
                    return;
                }

                if (!item.price || isNaN(item.price)) {
                    item.price = 0;
                }

                var pageNum = GetPageNum(item);

                var currentRow = info[pageNum].currentRow;
                var currentCol = info[pageNum].lastRowCurrentCol;

                //console.log(item.code);
                //console.log(item.gender);
                //console.log(currentRow + " " + currentCol);
                //console.log((currentRow + ItemImageRowSize) + " " + (currentCol + ItemImageColSize));

                var imgUrl = convertImage(__dirname + '/data/' + accessCode + '/' + item.url);

                sheets[pageNum].addImage({
                    path: imgUrl,
                    type: 'picture',
                    position: {
                        type: 'oneCellAnchor',
                        from: {
                            col: currentCol,
                            row: currentRow,
                            colOff: '0.5in'
                        }
                    }
                });
                sheets[pageNum]
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
                sheets[pageNum]
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
                sheets[pageNum]
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
                sheets[pageNum]
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

                sheets[pageNum]
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

                sheets[pageNum]
                    .cell(
                        currentRow + ItemImageRowSize + 7, currentCol,
                        currentRow + ItemImageRowSize + 7, currentCol+ ItemImageColSize - 1,
                        true
                    )
                    .style({
                        alignment: {
                            wrapText: true,
                            horizontal: 'center'
                        }
                    })
                    .string(item.price.toString());

                lastCol = currentCol+ ItemImageColSize - 1;
                info[pageNum].lastRowItemsCount++;
                info[pageNum].lastRowCurrentCol += ItemImageColSize;
            });

            sheets[pageNum]
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
                .string(fItem.name);

            sheets[pageNum]
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
                .string(fItem.code);

            sheets[pageNum]
                .cell(
                    fRow + ItemImageRowSize + 8, firstCol,
                    fRow + ItemImageRowSize + 8, lastCol,
                    true
                )
                .style({
                    alignment: {
                        wrapText: true,
                        horizontal: 'center'
                    }
                })
                .string(GetFormattedComments(fItem.code));


            //info[fItem.gender].currentCol ++;
            if (info[pageNum].lastRowItemsCount >= MaxPerRow) {
                info[pageNum].lastRowCurrentCol = 1;
                info[pageNum].lastRowItemsCount = 0;
                info[pageNum].currentRow += ItemImageRowSize + 13;
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

        wb.write('./data/' + accessCode + '/ExcelFile.xlsx', function(err, stats) {
            if (err) {
                console.log(err);
                return;
            }
            fs.readFile('./data/' + accessCode + '/ExcelFile.xlsx', function (err, data) {
                transporter.sendMail({
                    sender: 'testLacostApp@mail.ru',
                    to: email,
                    subject: name,
                    body: 'Order',
                    attachments: [{'filename': name + '.xlsx', 'content': data}]
                }, function (err, success) {
                    if (err) {
                        console.log(err);
                    }

                });
            });
        });
    }
};