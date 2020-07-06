import { Workbook, Cell  } from "exceljs";


 new Workbook().xlsx.readFile('E:\\Users\\Daniel\\Documents\\git working area\\test-excelj-cell-format\\test.xlsx').then((workbook:Workbook) => {
  const worksheet = workbook.getWorksheet("Sheet1")
  const cellgen = worksheet.getCell('B1')
  const celltext = worksheet.getCell('B2')
  const cellhhmm = worksheet.getCell('B3')
  const cellhhmmss = worksheet.getCell('B4')
  const cellhmmss = worksheet.getCell('B5')
  const cellhmm = worksheet.getCell('B6')

  console.log('Using General')
  console.log(getHourMinsFromMultiFormat(cellgen))
  console.log('Using Text')
  console.log(getHourMinsFromMultiFormat(celltext))
  console.log('Using hh:mm')
  console.log(getHourMinsFromMultiFormat(cellhhmm))
  console.log('Using hh:mm:ss')
  console.log(getHourMinsFromMultiFormat(cellhhmmss))
  console.log('Using [h]:mm:ss')
  console.log(getHourMinsFromMultiFormat(cellhmmss))
  console.log('Using [h]:mm]')
  console.log(getHourMinsFromMultiFormat(cellhmm))
  

 })

console.log("test")

const getHourMinsFromMultiFormat = (cell :Cell):{hours: number; mins: number } => {
  let startDate = convert('Dec 30 1899 00:00:00 GMT+0000 (Greenwich Mean Time)')

  switch (cell.numFmt) {
    case undefined: // General format
    if (/^\d+:\d+$/.test(cell.text)){
      let split = cell.text.split(":")
      return {hours: +split[0], mins: +split[1]}
    }
    if (/^\d+$/.test(cell.text)){
      return {hours: +cell.text, mins: 0}
    }

    return {hours: 0, mins: 0}
      break;
    case '@': // Text format
      if (/^\d+:\d+$/.test(cell.text)){
        let split = cell.text.split(":")
        return {hours: +split[0], mins: +split[1]}
      }
      if (/^\d+$/.test(cell.text)){
        return {hours: +cell.text, mins: 0}
      }

      return {hours: 0, mins: 0}

      break;
    case '0': // Text format
        return {hours: +cell.text, mins: 0}
      break;

    case 'hh:mm:ss':
      return hoursAndMinsBetweenTwoDates(startDate,convert(cell.text))
    break;
    
    case 'hh:mm':
      return hoursAndMinsBetweenTwoDates(startDate,convert(cell.text))
    break;
    
    case 'h:mm':
      return hoursAndMinsBetweenTwoDates(startDate,convert(cell.text))
    break;

    case 'h:mm:ss':
      return hoursAndMinsBetweenTwoDates(startDate,convert(cell.text))
    break;

    default:
      return {hours: 0, mins: 0}
  }
}

const convert = (date: string ): Date => {
  // Jan 09 1900 00:00:00 GMT+0000 (Greenwich Mean Time)
  let returnDate = new Date(date)
  return returnDate;
}


const hoursAndMinsBetweenTwoDates = (date1:Date, date2: Date): {hours: number; mins: number } => {
  let diff = (date2.getTime() - date1.getTime())
  let hours = Math.floor((diff/(1000*60*60)))
  let mins = Math.floor((diff/(1000*60)) % 60)
  let returnData = {hours: hours, mins: mins}


return returnData
}



