//This will send links out by email

const nssTheEmailer =(()=>{
  /**
   * All info for the email to be sent
   * @param {Object} emailInfo - The emailInfo object
   * @param {array} emailInfo.data - The data of all the sheetdata.
   * @param {string} emailInfo.recipient - The email of the recipient.
   * @param {string} emailInfo.subject - The subject for the email
   * @param {string} emailInfo.greeting - The opening of the message.
   * @param {string} emailInfo.closing - The closing of the message.
   */
  const sendTheMail_ = (emailInfo)=>{
    try{
      //get all the links into HTML rows   
      let theRows = emailInfo.data.map(row=>{
        return `<tr><td><a href=\"${row[1]}\">${row[2]}</a></td></tr>`
      }).join('')
      
      let temp = HtmlService.createTemplateFromFile("emailTemplate")
      temp.theRows = theRows
      temp.greeting = emailInfo.greeting
      temp.closing = emailInfo.closing
      let html = temp.evaluate().getContent().toString()
      //let msg= "<h1>This is big!!</h1><div style='color:purple;font-size:40px;'>This should be red.<div>"
      GmailApp.sendEmail(
        emailInfo.recipient,
        emailInfo.subject,
        "Body",
        {htmlBody:html}
      )
      // logIt({
      //   level:'info',
      //   theMsg:`Email sent to ${emailInfo.recipient}`
      // })
    }
    catch(err){
      logIt({
        level:"severe",
        theMsg:"Error in emailing links",
        error:err
      })
    }
    
  }
  return{
    sendTheMail:sendTheMail_
  }
})()

//This has to be AFTER the namespace and use 'var' so it will be "global"
var nsWEPemail = nssTheEmailer
