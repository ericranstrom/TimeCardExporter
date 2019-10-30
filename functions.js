var CATEGORIES = []
var ACCESS_TOKEN = ""

// The initialize function must be run each time a new page is loaded
Office.initialize = reason => {
    console.log('Office init...' + reason);
};

// Add any ui-less function here
function exportAppointments(event) {
    var buttonId = event.source.id;
    console.log('exportAppointments() called, buttonID: ' + buttonId);
    CATEGORIES = []
    var ACCESS_TOKEN = ""

    getAccessToken()
    .then(() => getIDsForWeeklyEvents())
    .then(function(idArr) {
        return getEventResponsesForIds(idArr)
    })
    .then(function(eventResponses){
        return getEventsForResponses(eventResponses)
    }).then(() => {
        console.log(CATEGORIES)
    })
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
        var getWeeklyEventsUrl = Office.context.mailbox.restUrl + "/v2.0/me/calendarview?startdatetime=2019-10-27T04:31:00.376Z&enddatetime=2019-11-03T04:31:00.376Z";
        console.log("make request for event ids with " + ACCESS_TOKEN)
        $.ajax({
            url: getWeeklyEventsUrl,
            dataType: 'json',
            headers: { 'Authorization': 'Bearer ' + ACCESS_TOKEN }
          }).done(function(response){
            console.log("Got the response from the rest api!")
            var ids = []
            response.value.forEach(function (item, index) {
                ids.push(item.Id)
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
              resolve(event)
            }).fail(function(error){
              reject(new Error("Failed to get event " + id))
            });
        });
        promises.push(p)
    });
    return Promise.all(promises)
}

function getEventsForResponses(eventResponses) {
    var events = []
    eventResponses.foreach(function(eventResponse, index) {
        if (eventResponse.Categories.length > 0) {
            var event = new Event(eventResponse.Subject, eventResponse.Start.DateTime, eventResponse.End.DateTime)
            newEvent(eventResponse.Categories[0], event)
        }
    });
}

function msToHumanReadable(ms) {
        var seconds = (ms/1000);
        var minutes = parseInt(seconds/60, 10);
        seconds = seconds%60;
        var quarterHours = parseInt(minutes/15, 10);
        minutes = minutes%15;
        var hours = parseInt(quarterHours/4, 10);
        quarterHours = quarterHours%4;

        return hours + '.' + (quarterHours * .25);
}



//Event Class Object
function Event(subject, startime, endtime) {
    this.subject = subject;
    this.durationInMillis = (new Date(endtime)).getTime() - (new Date(starttime)).getTime()
}

// Adding a method to the constructor
Event.prototype.greet = function() {
    return `${this.name} says hello.`;
}

function Category(name){
  this.name = name
  this.events = [];
}
Category.prototype = {
  constructor:Event,
  addEvent: function(event){
    this.events.push(event)
  }
};

function newEvent(categoryName, event){
  var catIndex = CATEGORIES.indexOf(categoryName)
  if (catIndex > -1) {
    CATEGORIES[catIndex].addEvent(event)
  } else {
    var category = new Category(categoryName)
    category.addEvent(event)
    CATEGORIES.push(category)
  }
}