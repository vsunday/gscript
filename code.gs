var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
var allEvents = []; //store event objects
//var calendarId = "calendarId@calendar.google.com";

function onOpen() {
  var menuEntries = [
    {
      name: "Update from Calendar",
      functionName: "syncFromCalendar"
    },
    {
      name: "Update to Calendar",
      functionName: "syncToCalendar"
    },
    {
      name: "Sync with Calendar",
      functionName: "syncWithCalendar"
    }
  ];
  spreadsheet.addMenu('Calendar Sync', menuEntries);
  prepareAllEvents();
}

function syncFromCalendar() {
  cal2Sheet()
}

function syncToCalendar() {
  sheet2Cal()
}

function sycWithCalendar() {
  sheet2Cal()
  cal2Sheet()
}

//store latest allEvents object in Sheet
function prepareAllEvents() {
  var sheet = spreadsheet.getActiveSheet();
  var range = sheet.getDataRange();
  var data = range.getValues();
  
  var titleRow = 1;
  
  var res = [];
  for (i=titleRow; i<data.length; i++) {
    var cal = new CalObj();
    cal.init(data[i])
    res.push(cal)
  }
  allEvents = res
}

Date.prototype.addDate = function(day) {
  this.setDate(this.getDate()+day);
  return this;
}

//add an event to calendar
function addEventOnCalendar(calObj) {
  var options = {
    description: calObj.summary,
    location: calObj.location
  }
  var calendar = CalendarApp.getCalendarById(calendarId);
  calendar.createAllDayEvent(calObj.title, new Date(calObj.startDate), (new Date(calObj.endDate)).addDate(1), options)
}

//update an existing event on calendar
function updateEventOnCalendar(existingEvent, newEvent) {
  existingEvent.setTitle(newEvent.title)
               .setLocation(newEvent.location)
               .setDescription(newEvent.summary)
  try {
    existingEvent.setAllDayDates(newEvent.startDate, (new Date(newEvent.endDate)).addDate(1))
  } catch(err) {
    existingEvent.setAllDayDate(newEvent.startDate)
  }
}

//get all events on calendar
function getEventsFromCalendar() {
  var calendar = CalendarApp.getCalendarById(calendarId);
  var term = {start:"1/1/2016", end:"12/31/2030"}
  var allCalEvents = calendar.getEvents(new Date(term.start), new Date(term.end));
  prepareAllEvents();
  for (i=0;i<allCalEvents.length;i++) {
    var cal = new CalObj();
    var flag = true;
    cal.loadFromCalendarEvent(allCalEvents[i]);
    for(j=0;j<allEvents.length;j++) {
      if ((cal.id==allEvents[j].id) || (allEvents[j].id == "" && allEvents[j].title == cal.title)) {
        allEvents[j].loadFromCalendarEvent(allCalEvents[i])
        flag = false;
        break;
      }
    }
    if (flag) allEvents.push(cal);
  }
}

//format method
function createRowFromCalObj(calObj) {
 return [calObj.id, calObj.title, calObj.startDate, calObj.endDate, calObj.summary, calObj.url, calObj.location]
}

//format method
function createRowsFromCalObjs(calObjs) {
  var res = [];
  for (i=0;i<calObjs.length;i++) {
    res.push(createRowFromCalObj(calObjs[i]))
  }
  return res
}

function sheet2Cal() {
  prepareAllEvents();
  var calendar = CalendarApp.getCalendarById(calendarId);
  for (i=0;i<allEvents.length;i++ ) {
    var sheetEvent = allEvents[i];
    if (!sheetEvent.id) {
      //add new event to calendar
      addEventOnCalendar(sheetEvent);
    } else {
     //update existing event on calendar 
      var existingEvent = calendar.getEventById(sheetEvent.id)
      updateEventOnCalendar(existingEvent, sheetEvent)
    }
  }
}

function cal2Sheet() {
  getEventsFromCalendar()
  var rangeText = "A2:G"+ (allEvents.length+1).toString()
  var range = spreadsheet.getActiveSheet().getRange(rangeText)
  range.setValues(createRowsFromCalObjs(allEvents))
}
