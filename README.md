# TimeCardExporter
Outlook Add-in To Export Meetings.

# Installation
 This application needs to be installed as a custom add-in to outlook.
* Open Outlook in the web client
* Open up an email or a calendar item
   * For an email, you may need to click the `...` near the reply button
   * For a calendar item, you may need to click the `...` near the delete button
* In the menu bar for the item, click 'Get Add-ins'
* In the menu panel on the left, select 'My Add-ins'
* Scroll down to the bottom section named 'Custom Add-ins'
* Click the button to 'Add a Custom Add-in'
    * Select 'Add from URL...'
    * Enter `https://ericranstrom.github.io/TimeCardExporter/manifest.xml`
    * Press OK

# Usage
* All of your meetings should be 'Categorized' into categories.
* Open up a calendar item that occurred during the week you wish to export a timecard for.
* In the menu ribbon, click the '...' and select `Time Card Exporter`
* This will download a CSV file containing the weekly timecard.

# Questions/Issues/Suggestions
* https://github.com/ericranstrom/TimeCardExporter/issues/new