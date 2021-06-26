const XLSX = require('xlsx');
const numberToString = require('number-to-cyrillic');
const { formatPrice, parseDate, getBuildingNo, getStreet } = require("./utils");
const moment = require('moment');
const { renderWrite } = require("./render-write");
moment.locale('ru');

//
const workbook = XLSX.readFile('input2.xlsx');

const first_sheet_name = workbook.SheetNames[0]; // sheet number
console.log('Processing first sheet with name ' + first_sheet_name);
const worksheet = workbook.Sheets[first_sheet_name];


const documents = [];

const start = 2;
let i = start;
while (worksheet['C' + i]) {
    const value = (letter, def = '') => worksheet[letter + i] && worksheet[letter + i].v || def;

    const address = value('C', null).trim();
    const serviceDates = parseDate(worksheet['B' + i]);


    serviceDates.forEach(serviceDate => {
        const existingAddress = documents.findIndex(document => document.address === address);

        const service = {
            service_date: serviceDate.format('DD.MM.YYYY'),
            service_name: value('H'),
            service_period: value('Q'),
            calculation_point: value('R') || '',
            smeta_price: formatPrice(value('P', 0)),
            service_price: formatPrice((value('O', 0)) / serviceDates.length),
            flat_no: value('D') || '',
        };

        if (existingAddress < 0) {
            documents.push({
                predsedatel: value('V'),
                flat_commander_no: value('W'),
                decision_title: value('X'),
                address: address,
                building_no: getBuildingNo(address),
                // flat_commander_no: value('M') || '',
                services: [service],
                services_total_price: service.service_price,
                service_total_price_text: numberToString.convert(service.service_price.toString(), { language: 'ru' }).convertedInteger,
                street: getStreet(address),
                act_date: serviceDate.format('MM/YYYY'),
                // datePeriodStartD: serviceDate.clone().startOf('month').format('DD'),
                // datePeriodStartM: serviceDate.format('MMMM'),
                // datePeriodStartY: serviceDate.year(),
                // datePeriodEndD: serviceDate.clone().endOf('month').format('DD'),
                // datePeriodEndM: serviceDate.format('MMMM'),
                // datePeriodEndY: serviceDate.year(),
            });
        } else {
            documents[existingAddress].services.push(service);
            documents[existingAddress].services_total_price = parseFloat(parseFloat(documents[existingAddress].services_total_price + service.service_price).toFixed(2));
            documents[existingAddress].service_total_price_text = numberToString.convert(documents[existingAddress].services_total_price.toString(), { language: 'ru' }).convertedInteger;
        }
    });

    i++;
}

console.log(`Parsing xls done. ${documents.length} addresses read`);


renderWrite(documents)
    .then(() => console.log('Done'))
;
