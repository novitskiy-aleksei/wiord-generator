const moment = require('moment');


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


//
function excelDateToJSDate(excel_date, time = false) {
    let day_time = excel_date % 1
    let meridiem = "AMPM"
    let hour = Math.floor(day_time * 24)
    let minute = Math.floor(Math.abs(day_time * 24 * 60) % 60)
    let second = Math.floor(Math.abs(day_time * 24 * 60 * 60) % 60)
    hour >= 12 ? meridiem = meridiem.slice(2, 4) : meridiem = meridiem.slice(0, 2)
    hour > 12 ? hour = hour - 12 : hour = hour
    hour = hour < 10 ? "0" + hour : hour
    minute = minute < 10 ? "0" + minute : minute
    second = second < 10 ? "0" + second : second
    let daytime = "" + hour + ":" + minute + ":" + second + " " + meridiem
    return time ? daytime : (new Date(0, 0, excel_date, 0, -new Date(0).getTimezoneOffset(), 0)).toLocaleDateString('ru', {}) + " " + daytime
}

module.exports = { getBuildingNo, getStreet, excelDateToJSDate };
