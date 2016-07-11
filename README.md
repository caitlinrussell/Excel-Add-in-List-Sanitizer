# Excel-Add-in-List-Sanitizer

This add-in sample demonstrates using the command ribbon to consolidate data in an Excel sheet. With the List Sanitizer, you can copy dump data from several sources (Mailchimp, server database, CRM, etc) and create one consolidated list of contacts.

Excel makes it easy to remove duplicate numerical data from a selection, but text consolidation requires a bit more Excel "magic". This plugin makes use of the Office.js API to grab and manipulate the sheet data in Javascript and return a single list in one click.

### Requirements
1. An Office365 account
2. Node, npm, and gulp installed

### Running the Sample
1. Clone or download the sample from Github
2. Inside the project folder, run 
```sh
$ gulp serve-static
```
The app is now running locally with HTTPS on port 8443
3. Load the manifest file in Excel
  1. In Excel Online:
    *. Click Insert > Office Add-ins
    *. Select "Upload My Add-in" in the Office Add-ins dialogue
  2. In Excel for Desktop:
    *. Share the project folder to a network share (\\\MyComputer\ListSantizer)
    *. In Excel, go to File > Options > Trust Center > Trusted Add-in Catalogs
    *. Add the URL of your network share and click "OK"
    *. Click Insert > My Add-ins
    *. Click on the Shared Folder tab and select your shared List Sanitizer app
The app will appear on the home ribbon

4. You can load your own contact list or use the demo Excel sheet contained in the project ("Customer Contact Lists.xlsx")
5. Make a selection and click "Sanitize Lists"!

![List Sanitizer screenshot](https://github.com/cbales/Excel-Add-in-List-Sanitizer/blob/master/readme-images/select-lists.png)

![Task pane screenshot](https://github.com/cbales/Excel-Add-in-List-Sanitizer/blob/master/readme-images/taskpane.png)

### About the Sample
This add-in was generated with the Office yo generator tool. Instructions on how to use the generator for Excel, Outlook, PowerPoint, and Word are at [dev.office.com](http://dev.office.com/getting-started/addins). You can also learn how to create a Visual Studio project from this page.

For a more detailed overview of the add-in process including writing a manifest file and accessing the API, take a look at the [office dev center docs](http://dev.office.com/docs/add-ins/develop/understanding-the-javascript-api-for-office)