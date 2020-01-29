let XLSX = require('xlsx');
let electron = require('electron').remote;
const parseString = require('xml2js').parseString;
const xml2js = require('xml2js');
const fs = require('fs');
let data = [];
let dataForObj = {};

let process_wb = (function () {
    let HTMLOUT = document.getElementById('htmlout');
    let XPORT = document.getElementById('xport');

    return function process_wb(wb) {
        XPORT.disabled = false;
        HTMLOUT.innerHTML = "";
        console.log('done');
        wb.SheetNames.forEach(function (sheetName) {
            HTMLOUT.innerHTML += XLSX.utils.sheet_to_html(wb.Sheets[sheetName], {editable: true});
            data = data.concat(XLSX.utils.sheet_to_json(wb.Sheets[sheetName]));
            //console.log(XLSX.utils.sheet_to_json(wb.Sheets[sheetName]));
        });
    };
})();

let xml = '';

let do_file = (function () {
    return function do_file(files) {
        let f = files[0];
        if (/.xml/.test(f.name)) {
            let parser = new xml2js.Parser();
            fs.readFile(f.path, function (err, data) {
                parser.parseString(data, function (err, result) {
                    console.dir(result);
                    console.log('Done');
                    xml = result;
                });
            });
        } else {
            let reader = new FileReader();
            reader.onload = function (e) {
                // noinspection JSUnresolvedVariable
                let data = e.target.result;
                data = new Uint8Array(data);
                process_wb(XLSX.read(data, {type: 'array'}));
            };
            reader.readAsArrayBuffer(f);
        }
    };
})();

(function () {
    let drop = document.getElementById('drop');

    function handleDrop(e) {
        e.stopPropagation();
        e.preventDefault();
        do_file(e.dataTransfer.files);
    }

    function handleDragover(e) {
        e.stopPropagation();
        e.preventDefault();
        e.dataTransfer.dropEffect = 'copy';
    }

    drop.addEventListener('dragenter', handleDragover, false);
    drop.addEventListener('dragover', handleDragover, false);
    drop.addEventListener('drop', handleDrop, false);
})();

(function () {
    let readf = document.getElementById('readf');

    function handleF(/*e*/) {
        let o = electron.dialog.showOpenDialog({
            title: 'Select a file',
            filters: [{
                name: "Spreadsheets",
                extensions: "xls|xlsx|xlsm|xlsb|xml|xlw|xlc|csv|txt|dif|sylk|slk|prn|ods|fods|uos|dbf|wks|123|wq1|qpw|htm|html".split(
                    "|")
            }],
            properties: ['openFile']
        });
        if (o.length > 0) process_wb(XLSX.readFile(o[0]));
    }

    readf.addEventListener('click', handleF, false);
})();

(function () {
    let xlf = document.getElementById('xlf');

    function handleFile(e) {
        do_file(e.target.files);
    }

    xlf.addEventListener('change', handleFile, false);
})();

let export_xlsx = (function () {
    let HTMLOUT = document.getElementById('htmlout');
    let XTENSION = "xls|xlsx|xlsm|xlsb|xml|csv|txt|dif|sylk|slk|prn|ods|fods|htm|html".split("|");
    return function () {
        let wb = XLSX.utils.table_to_book(HTMLOUT);
        let o = electron.dialog.showSaveDialog({
            title: 'Save file as',
            filters: [{
                name: "Spreadsheets",
                extensions: XTENSION
            }]
        });
        console.log(o);
        XLSX.writeFile(wb, o);
        electron.dialog.showMessageBox({message: "Exported data to " + o, buttons: ["OK"]});
    };
})();
void export_xlsx;

let builder = new xml2js.Builder();
let xmlOut = builder.buildObject(xml);

function exportFile() {
    console.log(dataForObj);
    fs.writeFile('data.js', 'let data = ' + JSON.stringify(dataForObj), (err) => {
        if (err) throw err;
    });
    console.log('The file has been saved!');
    /*builder = new xml2js.Builder();
    xmlOut = builder.buildObject(xml);
    fs.writeFile('pages.xml', xmlOut, (err) => {
        if (err) throw err;
        console.log('The file has been saved!');
    });*/
}

function makeObj() {
    dataForObj = data.map((item, index) => {
        let value = {};
        let temp = {};
        let check = {
            address: false,
            index: false,
            region: false,
            city: false
        };
        for (let key in item) {
            if (item.hasOwnProperty(key)) {
                if (key.trim() === 'Адреса п/в') {
                    check.address = true;
                    temp.address = item[key]
                }
                if (key.trim() === 'Область') {
                    check.region = true;
                    temp.region = item[key];
                }
                if (key.trim() === 'Населений пункт') {
                    check.city = true;
                    temp.city = item[key];
                }
                if (key.trim() === 'Індекс') {
                    check.index = true;
                    temp.index = item[key];
                }
                if (key.trim() === 'Індекс п/в') {
                    check.index = true;
                    temp.index = item[key];
                }
            }
        }
        if (check.address && check.region && check.city && check.index) {
            value = {
                label: temp.address + ' (' + temp.index + ')',
                value: temp.index,
                city: temp.region + ', ' + temp.city
            };
            return value
        }
        /*if (item['Адреса п/в'] !== undefined && item['Індекс '] !== undefined) {
            value = {
                label: item['Адреса п/в'] + ' (' + item['Індекс '] + ')',
                value: item['Індекс '],
                city: item['Область'] + ', ' + item['Населений пункт']
            };
            return value
        } else if (item['Адреса п/в'] !== undefined && item['Індекс п/в'] !== undefined) {
            value = {
                label: item['Адреса п/в'],
                value: item['Індекс п/в']
            };
            return value
        } else {
            console.log('We get some strange values on index: ', index, ' data: ', item);
        }*/
    }).filter(item => item !== undefined);
    console.log('total values: ', dataForObj.length);
}