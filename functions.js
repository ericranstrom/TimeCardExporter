// The initialize function must be run each time a new page is loaded
Office.initialize = reason => {
    console.log('Office init...' + reason);
};

// Add any ui-less function here
function exportAppointments(event) {
    var buttonId = event.source.id;
    console.log('exportAppointments() called, buttonID: ' + buttonId);



    Office.context.mailbox.getCallbackTokenAsync({isRest: true}, function(result){
      console.log(result)
      if (result.status === "succeeded") {
        var accessToken = result.value;
        console.log("access token: " + accessToken)

        // Use the access token.
        getCurrentItem(accessToken);
      } else {
        console.log("Failed to get access token!")
        // Handle the error.
      }
    });



    event.completed();
}

function getItemRestId() {
  if (Office.context.mailbox.diagnostics.hostName === 'OutlookIOS') {
    // itemId is already REST-formatted.
    return Office.context.mailbox.item.itemId;
  } else {
    // Convert to an item ID for API v2.0.
    return Office.context.mailbox.convertToRestId(
      Office.context.mailbox.item.itemId,
      Office.MailboxEnums.RestVersion.v2_0
    );
  }
}

function getCurrentItem(accessToken) {
  // Get the item's REST ID.
  var itemId = getItemRestId();

  // Construct the REST URL to the current item.
  // Details for formatting the URL can be found at
  // https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations#get-messages.
  var getMessageUrl = Office.context.mailbox.restUrl +
    '/v2.0/me/messages/' + itemId;

  $.ajax({
    url: getMessageUrl,
    dataType: 'json',
    headers: { 'Authorization': 'Bearer ' + accessToken }
  }).done(function(item){
    // Message is passed in `item`.
    var subject = item.Subject;
    console.log("Got the item from the rest api!")
    console.log(subject)
  }).fail(function(error){
    console.log("Failed to get item")
    // Handle error.
  });
}