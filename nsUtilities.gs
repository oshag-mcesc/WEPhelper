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
    const tableInfo = tables.map((table, index) => {
      // Check if the table has more than one row
      if (table.getNumRows() > 1) {
        // If it has more than one row, then check if the first row has at least 2 cells (columns)
        // We need to ensure the first row exists before trying to get its cells.
        // Although if getNumRows() > 1, getRow(0) should be safe.
        const firstRow = table.getRow(0);
        if (firstRow && firstRow.getNumCells() >= 2) {
          // If it meets both criteria (more than 1 row AND at least 2 columns),
          // return an array with data from specific cells and the table's original index.
          return [table.getCell(0, 1).getText(), table.getCell(1, 1).getText(), index];
        }
      }
      // If it doesn't meet the criteria, return "remove" to easily filter it out later.
      return "remove";
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

const check = ()=>{
  let id = "1TP9otI9qXdBarPLs0dnaSdfWBPXOZKXXcq3h49V1t74"
  let doc = DocumentApp.openById(id)
  let rslt = nsUtils.getTableIndexes(doc)
  console.log(rslt)
}