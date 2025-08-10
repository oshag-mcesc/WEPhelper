  /**
   * Info needed to log it through MyBBlog
   * @param {Object} info - The object with error message and error if there is one
   * @param {array} info.level - The level of the log message (ie severe, info, config)
   * @param {string} info.theMsg - The message to be logged
   * @param {string} info.error - The error object if there is one
   */
let logIt = (info)=>{
  let log = MyBBLog.getLog({
    sheetName:'Log1',
    level:'FINEST', 
    sheetId:'1UNOvucvbduwRzpgWPj0xxzBW7Knh-gkrgqORFuW2VHw',
    displayUserId:'EMAIL_FULL',
    });
  log[info.level](info.theMsg);
  if(info.error){
    log.severe(info.error.stack)
  }
  
}

//so scripts from sheet can log to error log sheet
var logError = logIt

const testLog=()=>{
      logIt({
        level:"info",
        theMsg:"Error in emailing links",
        //error:"Bad news!"
      })
}

var testLog1 = testLog
