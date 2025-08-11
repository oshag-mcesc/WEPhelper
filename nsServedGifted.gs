/**
 * Gets the students Gifted Areas and the course gifted area
 * It then sees if the course is in the students areas.  If so it is added to the list 
 * and the student will be "served" in that course.  If it is not then "NONE" is pushed to the list.
 *
 * Superior Cog and Creative Thinking are GIVENS (if a student has either or both and the class is a gifted class
 * then they are served!! EXCEPTIONS are for VPA!!)  EXCEPT for Huber Heights.  They do NO Supierior Cog AND Creative Thinking is done in Reading.
 * 
 * This all was updted 8/7/2025
 */

/**
 * @namespace
 */
const nsServedGifted = {};

/**
 * Main function to get the served gifted areas for students.
 * This acts as a router, calling the correct logic based on the school.
 * 
 * @param {boolean} huberElem - True if the school is a Huber Elem school.
 * @returns {boolean} True on success, false on failure.
 */
nsServedGifted.getServedGifted = (huberElem = false) => {
  try {
    const sheetName = 'preQuery';
    const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!ss) {
      logIt({
        level: 'severe',
        theMsg: `Could not find sheet named "${sheetName}".`,
        error: new Error(`Sheet not found: ${sheetName}`).stack
      });
      return false;
    }

    const lastRow = ss.getLastRow();
    if (lastRow < 2) {
      logIt({
        level: 'warning',
        theMsg: 'No data to process in the sheet.',
      });
      return true; // No data, but not an error
    }

    // HARD-CODED range
    const dataRange = ss.getRange(2, 11, lastRow - 1, 2);
    const allData = dataRange.getValues();

    let bigList;
    if (huberElem) {
      bigList = processHuberElem(allData);
    } else {
      bigList = processStandard(allData);
    }

    // HARD-CODED range
    ss.getRange('M1').setValue('Served');
    ss.getRange(2, 13, bigList.length, 1).setValues(bigList);

    return true;
  } catch (err) {
    logIt({
      level: 'severe',
      theMsg: 'Error getting list of served areas.',
      error: err.stack,
    });
    return false;
  }
};

/**
 * Processes the gifted data for a standard school.
 * @param {Array<Array<string>>} theData - The array of student data.
 * @returns {Array<Array<string>>} The list of served areas.
 */
const processStandard = (theData) => {
  const bigList = [];

  theData.forEach(data => {
    const gifted = [];
    // Need to watch out for VPA!! No SC or CT in VPA classes
    const studentGiftedAreas = data[0];
    const courseGiftedArea = data[1];

    if (studentGiftedAreas.indexOf('Superior Cognitive') >= 0 && courseGiftedArea.indexOf('Visual Performing Arts') < 0) {
      gifted.push(' Superior Cognitive');
    }

    if (studentGiftedAreas.indexOf('Creative Thinking') >= 0 && courseGiftedArea.indexOf('Visual Performing Arts') < 0) {
      gifted.push(' Creative Thinking');
    }

    if (studentGiftedAreas.indexOf(courseGiftedArea) >= 0) {
      gifted.push(` ${courseGiftedArea}`);
    }

    if (!gifted.length) {
      gifted.push('NONE');
    }

    bigList.push([gifted.toString()]);
  });
  return bigList;
};

/**
 * Processes the gifted data for a Huber Elem school.
 * @param {Array<Array<string>>} theData - The array of student data.
 * @returns {Array<Array<string>>} The list of served areas.
 */
const processHuberElem = (theData) => {
  const bigList = [];
  theData.forEach(data => {
    const gifted = [];
    const studentGiftedAreas = data[0];
    const courseGiftedArea = data[1];

    if (studentGiftedAreas.indexOf('Creative Thinking') >= 0 && courseGiftedArea === 'Reading') {
      gifted.push(' Creative Thinking');
    }

    if (studentGiftedAreas.indexOf(courseGiftedArea) >= 0) {
      gifted.push(` ${courseGiftedArea}`);
    }

    if (!gifted.length) {
      gifted.push('NONE');
    }

    bigList.push([gifted.toString()]);
  });
  return bigList;
};

/**
 * @namesapce
 * The following are functions that are called from the menu
 * to run the right logic for finding gifted area.
 */
var nsServedGiftedRouter = {}
//For standard logic
nsServedGiftedRouter.servedGiftedStandard = () => {
  nsServedGifted.getServedGifted()
}
//For HH to catch elem
nsServedGiftedRouter.servedGiftedHH = () => {
  let ui = SpreadsheetApp.getUi();
  let result = ui.prompt(
    'Set the flag!',
    'Please enter a 1 (or true) if you are dealing with a HH elementary school. \nEnter 0 (or false) for all others.',
    ui.ButtonSet.OK_CANCEL);
  // Process the user's response.
  let button = result.getSelectedButton();
  let text = result.getResponseText();
  let flag = 0;
  if (button == ui.Button.OK) {
    // User clicked "OK".
    flag = parseInt(text);
  } else if (button == ui.Button.CANCEL) {
    // User clicked "Cancel" get out!
    return;
  } else if (button == ui.Button.CLOSE) {
    // User clicked X in the title bar get out!
    return;
  }
  nsServedGifted.getServedGifted(flag)
}