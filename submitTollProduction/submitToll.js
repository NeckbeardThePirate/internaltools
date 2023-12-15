const XlsxPopulate = require('xlsx-populate');

async function submitTollCost(args) {
    const currentTollCost = args[0];
    const path = args[2];
    const user = args[1];
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
        const rowValues = [user, formattedDate,'Need to do this still',`Bridge Toll`,'Toll',currentTollCost]
        const letters = ['A','B','C','D','E','F','G','H','I'];
        for (let i = 0; i < letters.length; i++) {
            if (rowValues[i] !== undefined) {
                submissionsWorksheet.cell(`${letters[i]}${newRowNumber}`).value(rowValues[i])
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

const args = process.argv.slice(2);
submitTollCost(args)