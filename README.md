# Excel-Add-in-List-Sanitizer

This add-in sample demonstrates using the command ribbon to consolidate data in an Excel sheet. With List Sanitizer, you can copy dump data from several sources (Mailchimp, server database, CRM, etc) and create one consolidated list of contacts.

Excel makes it easy to remove duplicate numerical data from a selection, but text consolidation requires a bit more Excel "magic". This plugin makes use of the Office.js API to grab and manipulate the sheet data in Javascript and return a single list in one click.

### Requirements
1. An Office365 account
2. Node, npm, and gulp installed

### Running the Sample
1. Clone or download the sample from Github
2. Inside the project folder, run  ```$ gulp serve-static```
3. The app is now running locally with HTTPS on port 8443
4. Load the manifest file in Excel
  * In Excel Online:
    1. Click Insert > Office Add-ins
    2. Select "Upload My Add-in" in the Office Add-ins dialogue
  * In Excel for Desktop:
    1. Share the project folder to a network share (\\\MyComputer\ListSantizer)
    2. In Excel, go to File > Options > Trust Center > Trusted Add-in Catalogs
    3. Add the URL of your network share and click "OK"
    4. Click Insert > My Add-ins
    5. Click on the Shared Folder tab and select your shared List Sanitizer app
5. The app will appear on the home ribbon
6. You can load your own contact list or use the demo Excel sheet contained in the project ("Customer Contact Lists.xlsx")
7. Make a selection and click "Sanitize Lists"!

![List Sanitizer screenshot](https://github.com/cbales/Excel-Add-in-List-Sanitizer/blob/master/readme-images/select-lists.png)

![Task pane screenshot](https://github.com/cbales/Excel-Add-in-List-Sanitizer/blob/master/readme-images/taskpane.PNG)

### About the Sample
This add-in was generated with the Office yo generator tool. Instructions on how to use the generator for Excel, Outlook, PowerPoint, and Word are at [dev.office.com](http://dev.office.com/getting-started/addins). You can also learn how to create a Visual Studio project from this page.

For a more detailed overview of the add-in process, including writing a manifest file and accessing the API, take a look at the [office dev center docs](http://dev.office.com/docs/add-ins/develop/understanding-the-javascript-api-for-office).
