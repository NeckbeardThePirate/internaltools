const { response } = require("express");
const XlsxPopulate = require('xlsx-populate');

async function getTimecardsFromSyncro() {
    const APIToken = 'placeholder';
    const url = 'https://mcmahantech.syncromsp.com/api/v1/timelogs';
    const headers = new Headers();
    headers.append('Content-Type', 'application/json');
    headers.append('Authorization', APIToken);
    const today = new Date()
    const daysSinceLastSunday = today.getDay();
    const endDayOfPayPeriod = new Date();
    endDayOfPayPeriod.setDate(endDayOfPayPeriod.getDate() - (daysSinceLastSunday + 1));
    const startDayOfPayPeriod = new Date();
    startDayOfPayPeriod.setDate(startDayOfPayPeriod.getDate() - (7 + daysSinceLastSunday))
    const startMonth = startDayOfPayPeriod.getMonth() + 1;
    const endMonth = endDayOfPayPeriod.getMonth() + 1;
    const startDay = startDayOfPayPeriod.getDate();
    const endDay = endDayOfPayPeriod.getDate();
    try {
        fetch(url, {
            method: 'GET',
            headers: headers,
        })
        .then(response => response.json())
        .then(data => {
        let minutesWorked = 0;
            for (let i = 0; i < data.timelogs.length; i++) {
                const tempDateMonth = Number(data.timelogs[i].created_at.split('').slice(5,7).join(''));
                const tempDateDay = Number(data.timelogs[i].created_at.split('').slice(8,10).join(''));
                const monthIsSameAsStart =  startMonth <= tempDateMonth;
                const monthIsSameAsEnd = tempDateMonth <= endMonth;
                const dayIsAfterStart = (monthIsSameAsStart && startDay <= tempDateDay)
                const dayIsBeforeEnd = (monthIsSameAsEnd && tempDateDay <= endDay)
                if (dayIsAfterStart && dayIsBeforeEnd) {
                    const startTimeArray = data.timelogs[i].in_at.split('')
                    const endTimeArray = data.timelogs[i].out_at.split('')
                    const startTimeHours = Number(startTimeArray.slice(11, 13).join(''))
                    const endTimeHours = Number(endTimeArray.slice(11, 13).join(''))
                    const startMinutes = Number(startTimeArray.slice(14, 16).join(''))
                    const endMinutes = Number(endTimeArray.slice(14, 16).join(''))
                    const totalMinutesThisEntry = ((endTimeHours*60) + endMinutes) - ((startTimeHours*60) + startMinutes)
                    minutesWorked += totalMinutesThisEntry;
                }
            }
            console.log(minutesWorked)
            pushTimecardsIntoExcel(args, minutesWorked)
        })
        .catch(error => console.error('Error:', error));        
    } catch (error) {
        console.error('an error occured: ', error)
        console.warn('please do not contact McMahanTECH internal tools support')
        console.warn('this message will dissapear in 20 seconds')
        setTimeout(() => {}, 20000)
    }
}

async function pushTimecardsIntoExcel(args, minutesWorked) {
    const path = args[2];
    const user = args[1]
    try {
        const workbook = await XlsxPopulate.fromFileAsync(`${path}`)
        const submissionsWorksheet = workbook.sheet('AutomatedSubmissions')
        const newRowNumber = submissionsWorksheet.usedRange().endCell().rowNumber() + 1;
        const currentDate = new Date()
        const formattedDate = currentDate.toLocaleDateString('en-US', {
            year: 'numeric',
            month: '2-digit',
            day: '2-digit',
        });
        let timeWorked = minutesWorked/60;
        let OTWorked = 0;
        if (40 < timeWorked) {
            OTWorked += timeWorked - 40;
            timeWorked = 40;
        }
        const rowValues = [user, formattedDate,'Need to do this still',`Clocked Hours`,'Work Hrs.',timeWorked]
        const OTRowValues = [user, formattedDate,'Need to do this still',`Overtime Hours`,'OT Hrs.',OTWorked]
        const letters = ['A','B','C','D','E','F','G','H','I'];
        for (let i = 0; i < letters.length; i++) {
            if (rowValues[i] !== undefined) {
                submissionsWorksheet.cell(`${letters[i]}${newRowNumber}`).value(rowValues[i])
            }
            if (0 < OTWorked) {
                submissionsWorksheet.cell(`${letters[i]}${newRowNumber+1}`).value(OTRowValues[i])
            }
        }
        workbook.toFileAsync(`${path}`)
    } catch (error) {
        console.error('an error occured: ', error);
        console.warn('Please contact McMahan TECH internal tooling support')
        console.warn('This message will dissapear in 20 seconds')
        setTimeout(() => {}, 20000)
    }
}

// let args = process.argv.slice(2);
const args = ['4.5', 'Judah Helland', 'C:/Users/judah/Timecards_and_Reimbursements(copy).xlsx']

getTimecardsFromSyncro()
// pushTimecardsIntoExcel(args)

