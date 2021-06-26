const moment = require('moment');

/**
 * @param price string
 * @param min number
 * @param max number
 * @returns {number}
 */
const formatPrice = (price) => Math.round((price + Number.EPSILON) * 10000) / 10000;

function parseDate(dateRow) {
    if (dateRow.t === 'n') {
        return [moment(excelDateToJSDate(dateRow.v))];
    } else {
        const stringDate = dateRow.v.trim();
        const arrDates = stringDate.split(' ');
        return arrDates.filter(date => !!date).map(date => moment(date, 'DD.MM.YYYY'));
    }
}

function getBuildingNo(address) {
    let m;

    if ((m = /\d+\s?\D?\s?$/.exec(address)) !== null) {
        return m[0];
    } else {
        throw Error(`Failed to parse building # from address '${address}'`);
    }
}

function getStreet(address) {
    let m;

    if ((m = /^(.+)(\s|,)\d+\s?\D?\s?$/.exec(address)) !== null) {
        return m[1];
    } else {
        throw Error(`Failed to parse street from address '${address}'`);
    }
}


function excelDateToJSDate(serial) {
    const utc_days = Math.floor(serial - 25569);
    const utc_value = utc_days * 86400;
    const date_info = new Date(utc_value * 1000);

    const fractional_day = serial - Math.floor(serial) + 0.0000001;

    let total_seconds = Math.floor(86400 * fractional_day);

    const seconds = total_seconds % 60;

    total_seconds -= seconds;

    const hours = Math.floor(total_seconds / (60 * 60));
    const minutes = Math.floor(total_seconds / 60) % 60;

    return new Date(date_info.getFullYear(), date_info.getMonth(), date_info.getDate(), hours, minutes, seconds);
}

module.exports = { getBuildingNo, getStreet, excelDateToJSDate, formatPrice, parseDate };
