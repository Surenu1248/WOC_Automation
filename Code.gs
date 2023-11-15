function myFunction() {
  
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 1; 
  var numRows = sheet.getLastRow(), numColumns = sheet.getLastColumn(); 
  var dataRange = sheet.getRange(startRow, 1, numRows, sheet.getLastColumn());
  var data = dataRange.getValues(); 
  var satDay, sunDay;
  var weekEndDatesList=[], weekOffList=[], WOCNameList = [], WOCEmailList = [], finalResult = [];
  var presentWeekWOC = [], presentWeekWO = [];
  var totDays = 30;
  
  //console.log("Total Rows ",numRows, " ",sheet.getMaxColumns(), sheet.getLastColumn() );
  var matchNum = 0;

  //console.log(getWeekendList(data, numColumns, numRows))
  var outputList = getWeekendList(data, numColumns, numRows);

  console.log(getReceipientList(outputList[0], outputList[1]));

/*

getWeekendList(sheetData, noOfColumns, noOfRows);
This will retrun Weekend On Call Dates from the input sheet and filters who all scheduled
for the current weekend. It will also fetch all the Names / Emails information in the form a list.

*/
function getWeekendList(sheetData, noOfColumns, noOfRows)
{
  var currentWeekWOCList=[];
  for(var n=1; n<noOfColumns; n++)
  {
    //console.log(data);
      if(sheetData[1][n] == "Saturday" || sheetData[1][n] == "Sunday")
      {
        var currentWeekWOC = [];
        var tmpDate = new Date(sheetData[0][n]);
        tmpDate = Utilities.formatDate(tmpDate,'IST','MM/dd/yyyy');
        weekEndDatesList.push(tmpDate); // Total weekend dates in the current month
        currentWeekWOC.push(tmpDate);  // Current Week WOC Date
        //console.log('Weekend Day with Date ', data[0][n], " " ,data[1][n]);

        for(var rowsIter=2; rowsIter<noOfRows; rowsIter++)
        {
            if(sheetData[rowsIter][n] == "WOC")
            {
                currentWeekWOC.push(sheetData[rowsIter][0]); // Current Week WOC Name
                currentWeekWOC.push(sheetData[rowsIter][1]); // Current Week WOC Email
            }
        }
        currentWeekWOCList.push(currentWeekWOC);
        //finalResult.push(currentWeekWeekOffs);
      }
      else{
          for(var rowsIter1=2; rowsIter1<noOfRows; rowsIter1++)
            {
              if(sheetData[rowsIter1][n] == "WO")
              {
                var tmpList = []
                var tmpDate = new Date(sheetData[0][n]);
                tmpDate = Utilities.formatDate(tmpDate,'IST','MM/dd/yyyy');
                tmpList.push(tmpDate);
                tmpList.push(sheetData[rowsIter1][0]);
                tmpList.push(sheetData[rowsIter1][1]);
                weekOffList.push(tmpList);
                //console.log(currentWeekWeekOffs);
              }
            }
        }
  }
  //console.log(weekOffList);
  finalResult.push(currentWeekWOCList);
  finalResult.push(weekOffList);
  return finalResult;
}

/*
Get all recipient list to send an email reminder on Monday
*/
function getReceipientList(WOCOutput, WeekOffOutput)
{
  var tmpEmails=[],  wocDates = [];
  tmpNames = [];
  //console.log("===========WOC Output===========");
  //console.log(WOCOutput);
  //console.log("===========WO Output===========");
  //console.log(WeekOffOutput);
  
  for(var i=0; i<WOCOutput.length;i++)
  {
    var todayDate = new Date();
    var dayInNumber = todayDate.getDay();
    var WOCDate = new Date(WOCOutput[i][0]);
    
    var WOCDateMinus5 = new Date(WOCDate.getTime()-5*(24*3600*1000));
    WOCDateMinus5 = Utilities.formatDate(WOCDateMinus5,'IST','MM/dd/yyyy');
    //todayDate = new Date(todayDate.getTime()-5*(24*3600*1000));
    todayDate= Utilities.formatDate(todayDate,'IST','MM/dd/yyyy');
    //console.log(todayDate, WOCDateMinus5, dayInNumber);
    if(todayDate == WOCDateMinus5 && dayInNumber == 1) // Check if Monday +5 days logic to identify Sat and generate the Email List along with Woc dates
    {
      WOCOutput.forEach(function(val)
      {
        var tmpDate1 = Utilities.formatDate(WOCDate,'IST','MM/dd/yyyy'); // Get Saturday Date Format
        var tmpDate2 = new Date(WOCDate.getTime()+1*(24*3600*1000));
        tmpDate2 = Utilities.formatDate(tmpDate2,'IST','MM/dd/yyyy'); // Get Sundaty Date Format
        //console.log(val, tmpDate1, WOCDate , tmpDate2);
        if(tmpDate1 == val[0]  || tmpDate2 == val[0])
        {
          wocDates.push(val[0]);
          for(var j=2; j<val.length; j+=2)
          {
            tmpEmails.push(val[j]);
            tmpNames.push(val[j-1]);
          }
        }
        
      }
      );

      presentWeekWOC.push(WOCOutput[i]); // Complete data including Name, Date, Email
    }
  }
  //console.log(tmpEmails, tmpNames);
  //console.log(tmpNames);
  //console.log(wocDates);
  sendEmailToCurrentWOCEngineers(tmpEmails, wocDates, getWODates(WeekOffOutput, tmpNames));
  tmpEmails = tmpEmails.filter((item, 
        index) => tmpEmails.indexOf(item) === index);
  return tmpEmails;
}

// Get Current Month Week Off Dates based on The Names.
function getWODates(WeekOffOutput, inpNames)
{
    var currentWeekWODetails = [];
    //console.log("inputNames from Tmp Names " , inpNames);
    var dpNames = inpNames.filter((item, 
        index) => inpNames.indexOf(item) === index); 
    //console.log('After Dups', dpNames);
    WeekOffOutput.forEach(function(value)
    {
      for(var i=0; i <dpNames.length; i++)
      {
        if(dpNames[i] == value[1]  )
        {
          currentWeekWODetails.push(value[0]);
          currentWeekWODetails.push(dpNames[i]);
        }
      }
    }
    );
    var fnListForWODetails = [];
    dpNames.forEach(function(val2)
    {
      var tmpList = [];
      tmpList.push(val2);
      for(var il=1; il<currentWeekWODetails.length; il++)
      {
          if(val2 == currentWeekWODetails[il])
          {
            tmpList.push(currentWeekWODetails[il-1]);
          }
      }
      fnListForWODetails.push(tmpList); // WeekOffDates for Individual Persons
    }
    );
    //console.log("WODetails" , currentWeekWODetails);
    //console.log(fnListForWODetails);
    return fnListForWODetails;
}

// Send Email
function sendEmailToCurrentWOCEngineers(emailList, wocDates, woDates)
{
    var htmlTemplate = HtmlService.createTemplateFromFile("EmailOutput");

    console.log("=========== WOC Dates ===========");
    console.log(wocDates);
    console.log("======== Week Off Dates ==========");
    console.log(woDates);
    console.log("========= Email Details ==========");

    //============== HTML Mapings ================
    htmlTemplate.wocDate1 = wocDates[0];
    htmlTemplate.wocDate2 = wocDates[1];

    for(var i =0; i<woDates.length; i++)
    {
        htmlTemplate.woDate1 = woDates[i][1];
        htmlTemplate.woDate2 = woDates[i][2];
        var htmlforemail = htmlTemplate.evaluate().getContent();

        MailApp.sendEmail(emailList[i], "WOC Details", htmlforemail, {
          htmlBody:htmlforemail,
          cc:"",
          bcc: ""
          });
    }

    //var htmlforemail = htmlTemplate.evaluate().getContent();
    //MailApp.sendEmail("aansari@salesforce.com", "Test WOC Details", htmlforemail, {htmlBody:htmlforemail});

}

}







