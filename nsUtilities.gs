/**
 * Date Added: 11/17/2023
 * 
 * These utility functions are used by various scripts
 * 
 * TODO: add more info
 */
const nsUtils = (() => {
  /**
   * Utilty to get columns out of an array.  
   * @param {arr} theArray The array that is to be maniulated
   * @param {arr} colsArray The array of column indexes (0 based) that are to be KEPT
   * @return {arr} The array that is the result of the manipulation
   */
  const keepCols_ = (theArray, colsArray) => {

    let newArr = theArray.map(row => {
      return row.filter((el, idx) => {
        return colsArray.includes(idx)
      })
    })

    return newArr

  }
  //given a document it will index the tables with what is in 
  //the second column and both first and second row
  //it then adds the table index
  //returns array [[teacher, class, table index]]
  const getTableIndexes_ = (doc) => {
    //get all the tables
    let tables = doc.getBody().getTables()

    //map all the table info into a new array
    let tableInfo = tables.map((table, index) => {
      if (table.getNumRows() > 1) {
        return [table.getCell(0, 1).getText(), table.getCell(1, 1).getText(), index]
      } else {
        return "remove" //this makes it easy to remove tables from the final array by filtering
      }
    })

    //Filter the info to get rid of 'remove'
    let goodTableInfo = tableInfo.filter(row => {
      return (row !== "remove")
    })

    return goodTableInfo
  }
  /**
   * Utilty to convert milliseconds to days rounded to 2 decimals
   * @param {int} milliseconds The number of milliseconds to convert
   * @return {num} The number of days
   */
  const convertMillisecondsToDays_ = (milliseconds) => {
    return (milliseconds / (1000 * 60 * 60 * 24)).toFixed(2)
  }

  /**
   * Utilty to return the current date in time in the format of
   * MM/DD/YYY HH:MM:SS 
   * @return {date} The formated data
   */  
  const getFormattedDateTime_=() =>{
  const now = new Date();
  const options = {
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
  }

  const formattedDateTime = now.toLocaleString('en-US', options).replace(',', '');
  return formattedDateTime
}


  return {
    keepCols: keepCols_,
    getTableIndexes: getTableIndexes_,
    convertMillisecondsToDays: convertMillisecondsToDays_,
    getFormattedDateTime : getFormattedDateTime_
  }
})()