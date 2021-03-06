var CATEGORIES = []
var ACCESS_TOKEN = ""

// The initialize function must be run each time a new page is loaded
Office.initialize = reason => {
    console.log('TimeCardExporter initializing...' + reason);
};

// Add any ui-less function here
function exportAppointments(event) {
    var buttonId = event.source.id;
    CATEGORIES = []
    var ACCESS_TOKEN = ""

    getAccessToken()
    .then(() => getResponsesForWeeklyEvents())
    .then(eventResponses => makeEventsForResponses(eventResponses))
    .then(() => downloadCsv())
    .catch((error) => console.log("Error encountered while generating Time Card: " + error))
    .finally(() => event.completed());
}

function getAccessToken() {
    return new Promise((resolve, reject) => {
        Office.context.mailbox.getCallbackTokenAsync({isRest: true}, function(result){
            if (result.status === "succeeded") {
                ACCESS_TOKEN = result.value;
                resolve();
             } else {
                reject(new Error("Failed to get access token!"));
             }
           });
    });
}

//see response here https://docs.microsoft.com/en-us/previous-versions/office/office-365-api/api/version-2.0/calendar-rest-operations#GetEvent
function getResponsesForWeeklyEvents() {
    return new Promise((resolve, reject) => {
        var meetingDate = Office.context.mailbox.item.start;
        var priorSun = getLastSunday(meetingDate)
        var proceedingSat = getNextSaturday(meetingDate)
        var getWeeklyEventsUrl = Office.context.mailbox.restUrl + "/v2.0/me/calendarview?enddatetime=" + proceedingSat.toISOString() + "&startdatetime=" + priorSun.toISOString() + "&$select=Id,Subject,Categories,Start,End&$top=1000";
        $.ajax({
            url: getWeeklyEventsUrl,
            dataType: 'json',
            headers: { 'Authorization': 'Bearer ' + ACCESS_TOKEN, 'Prefer':'outlook.timezone="Pacific Standard Time"' }
          }).done(function(response){
            var items = []
            response.value.forEach(function (item, index) {
                items.push(item)
            });
            resolve(items)
          }).fail(function(error){
            reject(new Error("Failed to get weekly events"))
          });
    });
}

function makeEventsForResponses(eventResponses) {
    eventResponses.forEach(function(eventResponse, index) {
            if (eventResponse.Categories.length > 0) {
              newEvent(eventResponse.Categories[0], eventResponse.Subject, eventResponse.Start.DateTime, eventResponse.End.DateTime)
            }

    });
}

function downloadCsv() {
    const rows = []
    var headers = []
    headers.push('name')
    headers.push('memo')
    headers.push('sun')
    headers.push('mon')
    headers.push('tues')
    headers.push('wed')
    headers.push('thurs')
    headers.push('fri')
    headers.push('sat')
    rows.push(headers)
    console.log("Formatting " + CATEGORIES.length + " Categories for download")
    CATEGORIES.forEach(function (cat, id) {
        rows.push(cat.toList())
    });

    let csvContent = "data:text/csv;charset=utf-8,"
        + rows.map(e => e.join(",")).join("\n");

    var encodedUri = encodeURI(csvContent);
    var link = document.createElement("a");
    link.setAttribute("href", encodedUri);
    var meetingDate = Office.context.mailbox.item.start;
    var priorSun = getLastSunday(meetingDate)
    link.setAttribute("download", "week_of_"+ priorSun.getFullYear() + "_" + (priorSun.getMonth()+1) +"_"+priorSun.getDate()+".csv");
    document.body.appendChild(link); // Required for FF

    link.click(); // This will download the data file.

    window.open(encodedUri);

    console.log("The Document Should have downloaded")

    var pom = document.getElementById("download");
    var blob = new Blob([encodeURI(csvContent)],{type: 'text/csv;charset=utf-8;'});
    var url = URL.createObjectURL(blob);
    pom.href = url;
    pom.setAttribute('download', "week_of_"+ priorSun.getFullYear() + "_" + (priorSun.getMonth()+1) +"_"+priorSun.getDate()+".csv");
    pom.click();

    console.log("The Document Should have downloaded again")
}

function msToHumanReadable(ms) {
        var seconds = (ms/1000);
        var minutes = parseInt(seconds/60, 10);
        seconds = seconds%60;
        var quarterHours = parseInt(minutes/15, 10);
        minutes = minutes%15;
        var hours = parseInt(quarterHours/4, 10);
        quarterHours = quarterHours%4;

        return hours + (quarterHours * .25);
}



//Event Class Object
function Event(subject, starttime, endtime) {
    this.subject = subject;
    this.start = new Date(starttime);
    this.durationInMillis = (new Date(endtime)).getTime() - (new Date(starttime)).getTime()
}

// Adding a method to the constructor
Event.prototype = {
    constructor:Event,
    toString: function() {
        return 'subject: ' + this.subject + ' duration: ' + msToHumanReadable(this.durationInMillis)
    }
}

function Category(name){
  this.name = name;
  this.subjects = [];
  this.sun = 0;
  this.mon = 0;
  this.tues = 0;
  this.wed = 0;
  this.thurs = 0;
  this.fri = 0;
  this.sat = 0;
  this.events = [];
}

Category.prototype = {
  constructor:Category,
  addEvent: function(event){
    this.events.push(event)
  },
  toList: function() {
    //processEvents
    this.events.forEach(function(event, id){
        this.subjects.push(event.subject);
        switch(event.start.getDay()) {
          case 0: //sunday
            this.sun += event.durationInMillis;
            break;
          case 1: //monday
            this.mon += event.durationInMillis;
            break;
          case 2: //tuesday
            this.tues += event.durationInMillis;
            break;
          case 3: //wednesday
            this.wed += event.durationInMillis;
            break;
          case 4: //thursday
            this.thurs += event.durationInMillis;
            break;
          case 5: //friday
            this.fri += event.durationInMillis;
            break;
          case 6: //saturday
            this.sat += event.durationInMillis;
            break;
          default:
            // code block
        }
    }, this)

    row = []
    row.push(this.name)
    row.push(this.subjects.join(" | "))
    row.push(msToHumanReadable(this.sun))
    row.push(msToHumanReadable(this.mon))
    row.push(msToHumanReadable(this.tues))
    row.push(msToHumanReadable(this.wed))
    row.push(msToHumanReadable(this.thurs))
    row.push(msToHumanReadable(this.fri))
    row.push(msToHumanReadable(this.sat))
    return(row)
  },
  toString: function() {
    return '' + this.name + ',' + this.events;
  }
};

function newEvent(categoryName, subject, starttime, endtime){
  var event = new Event(subject, starttime, endtime)
  var category = getCategory(categoryName)
  if (category === undefined) {
     category = new Category(categoryName)
     category.addEvent(event)
     CATEGORIES.push(category)
  }
  else {
    category.addEvent(event)
  }
}

function getCategory(name){
  for (var i = 0; i < CATEGORIES.length; i+=1){
     if(CATEGORIES[i].name === name){
       return CATEGORIES[i]
     }
  }
}

function getLastSunday(d) {
  var t = new Date(d);
  t.setDate(t.getDate() - t.getDay());
  t.setHours(0,0,0,0);
  return t;
}

function getNextSaturday(d) {
  var t = new Date(d);
  t.setDate(t.getDate() + (6 - t.getDay()) % 7);
  t.setHours(23,59,59,999);
  return t;
}