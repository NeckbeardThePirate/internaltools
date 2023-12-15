const XlsxPopulate = require('xlsx-populate');

async function submitTripMileage(args) {
    console.log(args)
    setTimeout(() => {}, 12000)
    const user = args[0]
    const fromLocation = args[1]
    const toLocation = args[2];
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
    let miles;
    const path = args[3]; 
    try {
        const workbook = await XlsxPopulate.fromFileAsync(`${path}`)
        const tripMileageWorksheet = workbook.sheet("Trip Mileage")
        const submissionWorksheet = workbook.sheet("Submissions")
        const fromLocationPossibles = tripMileageWorksheet.find(fromLocation)
        for (const cell of fromLocationPossibles) {
            const rowNum = cell._row.rowNumber();
            const toLocationCellValue = tripMileageWorksheet.cell(`B${rowNum}`).value();
            if (toLocationCellValue === toLocation) {
                if (isRoundTrip) {
                    miles = tripMileageWorksheet.cell(`C${rowNum}`).value()
                    miles= miles*2
                } else {
                    miles = tripMileageWorksheet.cell(`C${rowNum}`).value()
                }
                break;
            }
        }
        const maxRow = submissionWorksheet.usedRange().endCell().rowNumber() + 1;
        const currentDate = new Date();
        const formattedDate = currentDate.toLocaleDateString('en-US', {
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
    });
        const values = [user, formattedDate,,,,`trip to ${toLocation} from ${fromLocation}`, 'Mileage',, miles]
        const letters = ['A','B','C','D','E','F','G','H','I',]
        for (let i = 0; i < values.length; i++) {
            submissionWorksheet.cell(`${letters[i]}${maxRow}`).value(values[i])
        }
        workbook.toFileAsync(`${path}`);   
    } catch (error) {
        console.error('an error occured: ', error);
        console.warn('Please do not contact McMahan TECH internal tooling support')
        console.warn('This message will dissapear in 20 seconds')
        setTimeout(() => {}, 20000)
    }
}


let args = process.argv.slice(2);

submitTripMileage(args)


//TODO 