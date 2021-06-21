const XLSX = require('xlsx');
const numberToString = require('number-to-cyrillic');
const { excelDateToJSDate, getBuildingNo, getStreet } = require("./utils");
const moment = require('moment');
moment.locale('ru');

//
const workbook = XLSX.readFile('input1.xlsx');

const first_sheet_name = workbook.SheetNames[0]; // sheet number
console.log('Processing first sheet with name ' + first_sheet_name);
const worksheet = workbook.Sheets[first_sheet_name];


const documents = [];

const start = 2;
let i = start;
while (true) {
    /* check if cell is empty in case of the end data */
    if (!worksheet['C' + i]) {
        break;
    }
    const value = (letter) => worksheet[letter + i] && worksheet[letter + i].v;

    const previousAddress = (0 < (i - start)) ? worksheet['C' + (i - 1)].v : null;
    const address = value('C');

    const service = {
        service_date: worksheet['B' + i].t === 'n' ? moment(excelDateToJSDate(value('B'))).format('DD.MM.YYYY') : worksheet['B' + i].v,
        service_name: value('H'),
        service_period: value('Q'),
        smeta_price: Math.floor(parseFloat(value('P'))),
        service_price: Math.floor(parseFloat(value('O') || 0)),
    };

    if (previousAddress !== address) {
        documents.push({
            customer: 'John Doe',
            address: address,
            building_no: getBuildingNo(address),
            flat_no: value('D') || '',
            flat_commander_no: '',
            // flat_commander_no: value('M') || '',
            services: [service],
            services_total_price: service.service_price,
            service_total_price_text: numberToString.convert(service.service_price.toString(), { language: 'ru' }).convertedInteger,
            street: getStreet(address),
            act_date: moment(service.service_date).format('MM/YYYY'),
            datePeriodStartD: moment(service.service_date).startOf('month').day(),
            datePeriodStartM: moment(service.service_date).startOf('month').format('MMMM'),
            datePeriodStartY: moment(service.service_date).startOf('month').year(),
            datePeriodEndD: moment(service.service_date).endOf('month').day(),
            datePeriodEndM: moment(service.service_date).endOf('month').format('MMMM'),
            datePeriodEndY: moment(service.service_date).endOf('month').year(),
        });
    } else {
        documents[documents.length - 1].services.push(service);
        documents[documents.length - 1].services_total_price += service.service_price;
        documents[documents.length - 1].service_total_price_text = numberToString.convert(documents[documents.length - 1].services_total_price.toString(), { language: 'ru' }).convertedInteger;
    }

    i++;
}

// documents {
//  address
//  street
//  ...
//  services: [
//    service_date
//    ...
//  ]
// }

console.log(`Parsing xls done. ${documents.length} addresses read`);


// write to docx
// ============================

const fs = require('fs');
const path = require('path');
const Docxtemplater = require("docxtemplater");
const PizZip = require("pizzip");
//


documents.forEach(document => {
    //Load the docx file as a binary
    const content = fs.readFileSync(path.resolve(__dirname, 'template.docx'), 'binary');
    const zip = new PizZip(content);
    let doc;
    try {
        doc = new Docxtemplater(zip);
    } catch(error) {
        // Catch compilation errors (errors caused by the compilation of the template : misplaced tags)
        console.error(error);
    }

    //set the templateVariables
    doc.setData(document);
    //

    try {
        // render the document (replace all occurences of {first_name} by John, {last_name} by Doe, ...)
        doc.render();
    } catch (error) {
        // Catch rendering errors (errors relating to the rendering of the template : angularParser throws an error)
        console.error(error);
    }

    const buf = doc.getZip().generate({type: 'nodebuffer'});
    // buf is a nodejs buffer, you can either write it to a file or do anything else with it.
    fs.writeFileSync(path.resolve(__dirname, `output/${document.address}.docx`), buf);

    throw new Error('ok');
});

console.log('Done');
