var xlsx = require('node-xlsx').default;
const fs = require('fs');
const moment = require('moment');


const options = {
    type: "buffer",
    cellDates: true,
    cellHTML: false,
    dateNF: "yyyy-mm-dd hh:mm:ss:s",
    cellNF: true
};
// Parse a file
const workSheetsFromFile = xlsx.parse(`${__dirname}/myFile.xlsx`,options);

console.log(workSheetsFromFile[0].data.length);

const updateQuery = 'update tms_Transaction_ETC set \n' +
    'TransactionDateTime=\'DATE_TIME\',\n' +
    'TagReadDateTime=\'DATE_TIME \'\n' +
    ' where CreatedOn>\'2022-06-28 00:00:00\'\n' +
    'and CreatedOn<\'2022-06-30 00:00:00\' \n' +
    'and Attribute_1=\'AA Device\'\n' +
    'and VehicleNumber=\'VEHICLE_NUMBER\';';

workSheetsFromFile[0].data.map(e =>{
    const q = updateQuery;
    console.log(e[2],'$$$$');
    const parsedDate = moment(e[2]).format('YYYY-MM-DD HH:mm:s');
    const q1 = q.replace('VEHICLE_NUMBER',e[5]);
    const q2 = q1.replaceAll('DATE_TIME',parsedDate)
    fs.appendFile('sqlOutput.txt', `${q2}\n\n`, (err) => {
        if (err) throw err;
        // console.log('updated.');
    });
})

