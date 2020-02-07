/**
 * Gets the students Gifted Areas and the course gifted area
 * It then sees if the course is in the students areas.  If so it is added to the list 
 * and the student will be "served" in that course.  If it is not then "NONE" is pushed to the list.
 * 
 *FBPS & FBIS - SC All, CT only w/ Reading, NO VPA
 *FBMS - SC non VPA, CT with Math and Reading EXCEPT Math 8 and ELA7-purple
 * 
 * @param {obj} theSheet the sheet with all the data
 * @param {array} theData the array of data values
 * @param {string} school maybe used for specifc school
 * @param {array} extraDataMS (optional) Need course data for MS to check CT Math8 and ELA7purple the array of data values
 * @return {boolean} true when done
 */

function getServedGiftedFairborn(theSheet, theData, school, extraDataMS) {
  var rslts;
  switch(school){
    case "PS":
    case "IS":
      rslts = servedPSandIS_(theData);
      break;
    case "MS":
      rslts = servedMS_(theData, extraDataMS);
      break;
   case "HS":
      rslts = servedHS_(theData);
      break;
  }
  if(rslts.length>1){ 
     //HARD CODED range!!!!!!!  ****************
    theSheet.getRange("L1").setValue("Served");
    theSheet.getRange(2,12,theSheet.getLastRow(),1).setValues(rslts);
    return true;
    }
  else{
    return false;
    }
}

function servedPSandIS_(theData){
  var bigList = [];
  var theList = '';
  
  theData.forEach(function(data){
    var gifted =[];
    //SC is a given!!
    if(data[0].indexOf("Superior Cognitive")>=0){
        gifted.push(' Superior Cognitive');
    }
    //CT only if in Reading
    if(data[0].indexOf("Creative Thinking")>=0 && data[1]=='Reading'){
    gifted.push(' Creative Thinking');
    }
    //Do the others now
    if(data[0].indexOf(data[1])>=0){
    gifted.push(" " + data[1]);}
    
    if(!gifted.length){gifted.push("NONE");}
    
    bigList.push([gifted.toString()]);
    });  
    
  return bigList;

}

function servedMS_(theData, extraDataMS) {
  var bigList = [];
  var theList = '';
  
  theData.forEach(function(data, index){
    var gifted =[];
    //Need to watch out for VPA!! No SC  in VPA classes
    if(data[0].indexOf("Superior Cognitive")>=0){
      if(data[1].indexOf("Visual Performing Arts")<0){
        gifted.push(' Superior Cognitive');
      }
    }
    //If CT then we have some hoops to jump through!!!!  
    if(data[0].indexOf("Creative Thinking")>=0){
      switch(data[1]){
        case "Math": //If not Math 8...
          if(extraDataMS[index].indexOf("Math 8")<0){
            gifted.push(' Creative Thinking');
            }
          break;
        case "Reading": //If not ELA7 purple...
          if(extraDataMS[index].indexOf("English/Language Arts 7-Purple")<0){
            gifted.push(' Creative Thinking');
            }
          break;      
        case "Visual Performing Arts": //If VPA...
            gifted.push(' Creative Thinking');
          break;     
      }//close the switch
    }//close the if
    //do the others now
    if(data[0].indexOf(data[1])>=0){
    gifted.push(" " + data[1]);}
    
    if(!gifted.length){gifted.push("NONE");}
    
    bigList.push([gifted.toString()]);
    });  
    
  return bigList;
}

function servedHS_(theData){
  var bigList = [];
  var theList = '';
  
  theData.forEach(function(data){
    var gifted =[];
    //SC with NO VPA
    if(data[0].indexOf("Superior Cognitive")>=0){
      if(data[1].indexOf("Visual Performing Arts")<0){
        gifted.push(' Superior Cognitive');
      }
    }    
    
    //CT is a given!!
    if(data[0].indexOf("Creative Thinking")>=0){
        gifted.push(' Creative Thinking');
    }

    //Do the others now
    if(data[0].indexOf(data[1])>=0){
    gifted.push(" " + data[1]);}
    
    if(!gifted.length){gifted.push("NONE");}
    
    bigList.push([gifted.toString()]);
    });  
    
  return bigList;

}