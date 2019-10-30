var CATEGORIES = []
var ACCESS_TOKEN = ""

// The initialize function must be run each time a new page is loaded
Office.initialize = reason => {
    console.log('Office init...' + reason);
};

// Add any ui-less function here
function exportAppointments(event) {
    var buttonId = event.source.id;
    CATEGORIES = []
    var ACCESS_TOKEN = ""

    getAccessToken()
    .then(() => getIDsForWeeklyEvents())
    .then(idArr => getEventResponsesForIds(idArr))
    .then(eventResponses => makeEventsForResponses(eventResponses))
    .then(() => CATEGORIES.forEach(function(cat, idx) {console.log(cat)}))
    .then(() => downloadCsv())
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

function getIDsForWeeklyEvents() {
    return new Promise((resolve, reject) => {
        var meetingDate = Office.context.mailbox.item.start;
        console.log(meetingDate)
        var priorSun = getLastSunday(meetingDate)
        console.log(priorSun)
        var proceedingSun = getNextSunday(meetingDate)
        console.log(proceedingSun)
        //var getWeeklyEventsUrl = Office.context.mailbox.restUrl + "/v2.0/me/calendarview?startdatetime=2019-10-27T04:31:00.376Z&enddatetime=2019-11-03T04:31:00.376Z";
        var getWeeklyEventsUrl = Office.context.mailbox.restUrl + "/v2.0/me/calendarview?enddatetime=" + proceedingSun.toISOString() + "&startdatetime=" + priorSun.toISOString() + "&$select=Id,Subject,Categories,Start,End";
        $.ajax({
            url: getWeeklyEventsUrl,
            dataType: 'json',
            headers: { 'Authorization': 'Bearer ' + ACCESS_TOKEN }
          }).done(function(response){
            var ids = []
            console.log(response.value)
            response.value.forEach(function (item, index) {
                ids.push(item.Id)
                console.log(item.Subject)
            });
            console.log(ids)
            resolve(ids)
          }).fail(function(error){
            reject(new Error("Failed to get weekly events"))
          });
    });
}

//see response here https://docs.microsoft.com/en-us/previous-versions/office/office-365-api/api/version-2.0/calendar-rest-operations#GetEvent
function getEventResponsesForIds(ids) {
    var promises = []
    ids.forEach(function (id, index) {
        var getEventUrl = Office.context.mailbox.restUrl + "/v2.0/me/events/"+id;
        var p = new Promise((resolve, reject) => {
          $.ajax({
              url: getEventUrl,
              dataType: 'json',
              headers: { 'Authorization': 'Bearer ' + ACCESS_TOKEN }
            }).done(function(response){
              console.log("Got the event response from the rest api!")
              console.log(response.Subject)
              console.log(response.Categories)
              console.log(response.Start.DateTime)
              console.log(response.End.DateTime)
              var meetingMillis = (new Date(response.End.DateTime)).getTime() - (new Date(response.Start.DateTime)).getTime()
              console.log(meetingMillis)
              console.log(msToHumanReadable(meetingMillis))
              console.log("*********************************")
              resolve(response)
            }).fail(function(error){
              reject(new Error("Failed to get event " + id))
            });
        });
        promises.push(p)
    });
    return Promise.all(promises)
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
    CATEGORIES.forEach(function (cat, id) {
        row = []
        row.push(cat.name)
        row.push(cat.getHours())
        rows.push(row)
    });

    let csvContent = "data:text/csv;charset=utf-8,"
        + rows.map(e => e.join(",")).join("\n");

    var encodedUri = encodeURI(csvContent);
    var link = document.createElement("a");
    link.setAttribute("href", encodedUri);
    link.setAttribute("download", "my_data.csv");
    document.body.appendChild(link); // Required for FF

    link.click(); // This will download the data file named "my_data.csv".
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
  this.name = name
  this.events = [];
}
Category.prototype = {
  constructor:Category,
  addEvent: function(event){
    this.events.push(event)
  },
  getHours: function() {
    var sumOfDuration = 0
    this.events.forEach(function(event, id){
      sumOfDuration += event.durationInMillis
    })
    return msToHumanReadable(sumOfDuration)
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
  return t;
}

function getNextSunday(d) {
  var t = new Date(d);
  t.setDate(t.getDate() + (7 - t.getDay()) % 7);
  return t;
}