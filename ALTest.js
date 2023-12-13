const ExcelJS = require('exceljs')
const fs = require('fs');
function logTripMileage(args) {

    const user = args[2]
    const fromLocationArg = args[0]
    const toLocationArg = args[1];
    let isRoundTrip = true;
    if (4 < args.length) {
        if (args[4] === 'RT') {
            isRoundTrip = true;
        }
        if (args[4] === 'OW') {
            isRoundTrip = false;
        }
    } else {
        isRoundTrip = true;
    }

    const doc = new ExcelJS.Workbook();
    let miles;
    
    doc.xlsx.readFile("C:/Users/judah/Timecards and Reimbursements(copy).xlsx")
        .then(() => {

            // console.log(doc.worksheets)
            let num = 0;
            for (sheet in doc.worksheets) {
                num++
            }
            console.log(num)
            tripMileageWorksheet = doc.getWorksheet(4);


            submissionsWorksheet = doc.getWorksheet(1);

            const fromLocationColumn = tripMileageWorksheet.getColumn('A');
            let validationCell;

            fromLocationColumn.eachCell(function(cell, rowNumber) {
                if (cell.value === fromLocationArg) {
                    if (isRoundTrip) {
                        validationCell = tripMileageWorksheet.getCell(`D${rowNumber}`)
                    } else {
                        validationCell = tripMileageWorksheet.getCell(`C${rowNumber}`)
                    }
                    miles = validationCell.value
                }

            });

            const nameColumn = submissionsWorksheet.getColumn('A')
            const rowCount = submissionsWorksheet._rows;
            console.log(rowCount.length)
            const date = Date.now() //still need to format this date

            // const newRow = submissionsWorksheet.addRow([user, date,,,, `Trip from ${fromLocationArg} to ${toLocationArg}`, 'Mileage',, `${miles}`])
            // console.log(newRow)
            console.log(rowCount.length)

            console.log(submissionsWorksheet.name)

        }).catch((err) => {
            console.error("error reading file: ", err)
        });
    return doc.xlsx.writeBuffer()
    .then(buffer => {
        fs.writeFileSync("C:/Users/judah/Timecards and Reimbursements(copy).xlsx", buffer);
    })
    .catch((err) => {
        console.error("Error writing file:", err);
    });

    return

}

// let args = process.argv.slice(2);
let args = ['MCTHQ', 'test', 'testuser', `trip to `]

logTripMileage(args);
