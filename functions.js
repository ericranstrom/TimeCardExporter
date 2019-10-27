// The initialize function must be run each time a new page is loaded
Office.initialize = reason => {
    console.log('Office init...' + reason);
};

// Add any ui-less function here
function exportAppointments(event) {
    var buttonId = event.source.id;
    console.log('exportAppointments() called, buttonID: ' + buttonId);

    Office.context.mailbox.getCallbackTokenAsync({isRest: true}, function(result){
      if (result.status === "succeeded") {
        var accessToken = result.value;
        // Use the access token.
        getWeeklyEvents(accessToken, function() {
          event.completed();
        });
      } else {
        console.log("Failed to get access token!")
        // Handle the error.
      }
    });
}

function getWeeklyEvents(accessToken, callback) {

  var getWeeklyEventsUrl = Office.context.mailbox.restUrl + "/v2.0/me/calendarview?startdatetime=2019-10-27T04:31:00.376Z&enddatetime=2019-11-03T04:31:00.376Z";

  $.ajax({
    url: getWeeklyEventsUrl,
    dataType: 'json',
    headers: { 'Authorization': 'Bearer ' + accessToken }
  }).done(function(response){
    console.log("Got the response from the rest api!")

    response.value.forEach(function (item, index) {
      getEventItem(accessToken, item.Id)
    });

    console.log(response.value[0])
    callback()
  }).fail(function(error){
    console.log("Failed to get item")
    // Handle error.
    callback()
  });
}

function getEventItem(accessToken, id) {
  var getEventUrl = Office.context.mailbox.restUrl + "/v2.0/me/events/"+id;
  $.ajax({
    url: getEventUrl,
    dataType: 'json',
    headers: { 'Authorization': 'Bearer ' + accessToken },
    async: false
  }).done(function(response){
    console.log("Got the event response from the rest api!")
    console.log(response)
    console.log(response.Subject)
  }).fail(function(error){
    console.log("Failed to get event")
  });

}