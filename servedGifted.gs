/**
 * Gets the students Gifted Areas and the course gifted area
 * It then sees if the course is in the students areas.  If so it is added to the list 
 * and the student will be "served" in that course.  If it is not then "NONE" is pushed to the list.
 *
 * Superior Cog and Creative Thinking are GIVENS (if a student has either or both and the class is a gifted class
 * then they are served!!)  EXCEPT for Huber Heights.  They do NO Supierior Cog AND Creative Thinking is done in Reading.
 * 
 * @param {obj} theSheet the sheet with all the data
 * @param {array} theData the array of data values
 * @param {bool} huberElem if this is a Huber Elem School then this is true
 * @return {boolean} true when done
 */

function getServedGifted(theSheet, theData, huberElem) {
  //huberElem is needed because they do no Superioar Cog...
  if(huberElem ==='undefined'|| huberElem === null){huberElem = false;}
  
  var bigList = [];
  var theList = '';
  
  if(!huberElem){
  theData.forEach(function(data){
    var gifted =[];
    if(data[0].indexOf("Superior Cognitive")>=0){
    gifted.push(' Superior Cognitive');}
    if(data[0].indexOf("Creative Thinking")>=0){
    gifted.push(' Creative Thinking');}
    
    if(data[0].indexOf(data[1])>=0){
    gifted.push(" " + data[1]);}
    
    if(!gifted.length){gifted.push("NONE");}
    
    bigList.push([gifted.toString()]);
    });  
  }
  else{
  theData.forEach(function(data){
    var gifted =[];
    if(data[0].indexOf("Creative Thinking")>=0 && data[1]=='Reading'){
    gifted.push(' Creative Thinking');}
    
    if(data[0].indexOf(data[1])>=0){
    gifted.push(" " + data[1]);}
    
    if(!gifted.length){gifted.push("NONE");}
    
    bigList.push([gifted.toString()]);
    });  
  }
  
   //HARD CODED range!!!!!!!  ****************
  theSheet.getRange(2,12,theSheet.getLastRow(),1).setValues(bigList);
  
  return true;
}