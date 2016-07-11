(function(){
  'use strict';

  // The initialize function must be run each time a new page is loaded
  Office.initialize = function(reason){
    jQuery(document).ready(function(){
    });
  };

})();

// Reads data from current document selection and condense it into one column
function sanitizeSelection(){

   Excel.run(function (ctx){
      var selectedRange = ctx.workbook.getSelectedRange();
      selectedRange.load('text');

      return ctx.sync().then(function(){
        var combinedList = [];

        //We want to display these vertically instead of horizontally
        var colList = [];

       for(var i = 0; i<selectedRange.text.length; i++) {
          for (var j = 0; j<selectedRange.text[i].length; j++) {
            var lineItem = selectedRange.text[i][j];
            if($.inArray(lineItem, combinedList) === -1 && lineItem.length > 0) {
              combinedList.push(lineItem);
              colList.push([lineItem]);
            }
          }
        }

        colList.sort();
        

        //Put the results in a new spreadsheet
        var worksheetCollection = ctx.workbook.worksheets;
        var worksheet = worksheetCollection.add();

        worksheet.load('text');
        return ctx.sync().then(function(){
          worksheet.activate();

          //"Paste" the results into the first column
          var rangeAddress = "A1:A"+colList.length;
          var range = worksheet.getRange(rangeAddress);
          range.values = colList;
        });
      });
   });
}