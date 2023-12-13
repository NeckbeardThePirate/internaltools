const ExcelJS = require('exceljs')

const str = 'it works';

const doc = new ExcelJS.Workbook();

doc.xlsx.readFile('C:/Users/judah/test00.xlsx')
    .then(() => {
        worksheet = doc.getWorksheet(1);

        const cellB2 = worksheet.getCell('B2');
        cellB2.value = str;

        return doc.xlsx.writeFile('C:/Users/judah/test00.xlsx');
    })

// doc.xlsx.readFile('C:/Users/judah/OneDrive - McMahan TECH LLC/Documents - Operations/Expenses/Timecards and Reimbursements.xlsx')
//     .then(() => {
//         worksheet = doc.getWorksheet(1);
//         // console.log(worksheet)

//         const theVar = worksheet.getCell('I10');
//         // // cellB2.value = str;
//         // console.log(worksheet.name)
//         console.log(theVar.value)
//         // for (column in worksheet._columns) {
//         //     // if (cell.value === 'it works') {
//         //     //     console.log('this is THE ONE', cell.address)
//         //     // }
//         //     // else {
//         //     //     console.log(cell.address)
//         //     // }
//         //     if (column == 0) {
//         //         console.log(worksheet._rows)
//         //     }
//         // }

//         return doc.xlsx.writeFile('C:/Users/judah/test00.xlsx');
//     })



    // let args = process.argv.slice(2);