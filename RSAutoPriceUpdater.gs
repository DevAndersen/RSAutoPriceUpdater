//The name of the sheet this script uses. If you need it to be called something else, change this variable to something else.
var priceSheetName = "Prices";

//Adds the menu to the spreadsheet
function onOpen()
{
  SpreadsheetApp.getUi().createMenu("RuneScape Price Updater")
      .addItem("Manual", "showManualPrompt_")
      .addItem("Generate price sheet", "generatePriceUpdateSheet")
      .addSeparator()
      .addItem("Update price for row...", "showUpdateSpecificRowPrompt_")
      .addSeparator()
      .addItem("Update prices for row 2 - 49", "updateRowsBatch1")
      .addItem("Update prices for row 50 - 99", "updateRowsBatch2")
      .addItem("Update prices for row 100 - 149", "updateRowsBatch3")
      .addItem("Update prices for row 150 - 199", "updateRowsBatch4")
      .addItem("Update prices for row 200 - 249", "updateRowsBatch5")
      .addItem("Update prices for row 250 - 299", "updateRowsBatch6")
      .addItem("Update prices for row 300 - 349", "updateRowsBatch7")
      .addItem("Update prices for row 350 - 399", "updateRowsBatch8")
      .addItem("Update prices for row 400 - 449", "updateRowsBatch9")
      .addItem("Update prices for row 450 - 499", "updateRowsBatch10")
      .addSeparator()
      .addItem("About", "showAboutPrompt_")
      .addToUi();
}

//Display prompt  displayed when updating a specific row
function showUpdateSpecificRowPrompt_()
{
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt("Update price for a specific row", "Row", ui.ButtonSet.OK_CANCEL);
  var button = result.getSelectedButton();
  var text = result.getResponseText();
  
  if (button == ui.Button.OK)
  {
    if(isNumber_(text))
    {
      updatePricesForRange_(text, text);
    }
    else
    {
      ui.alert("'" + text + "' is not a valid number.");
    }
  }
}

function showManualPrompt_()
{
  var textDescription = "This script will generate a sheet which will automatically keep item prices up to date. Click the 'Generate price sheet' button in the RuneScape Price Updater dropdown menu to get started.";
  textDescription += "\nThen, simply add the names and IDs of the items you wish to track to the " + priceSheetName + " sheet, and the script will handle the rest.";
  var textHowToRefer = "To get the price of an item into another sheet, simply set the text of cell you want to fill with '=" + priceSheetName + "!D2'.";
  textHowToRefer += " Replace the number 2 to reflect the row that corresponds to the item of interest.";
  var textItemID = "To get the ID of an item:";
  textItemID += "\n- Go to the RuneScape homepage (www.runescape.com)";
  textItemID += "\n- Go to the Grand Exchange part of the website";
  textItemID += "\n- Search for the item you're interested in, and go to that items' page";
  textItemID += "\n- The item's ID will be the last part of the URL";
  textItemID += "\n- For example, the Cabbage URL is 'http://services.runescape.com/m=itemdb_rs/Cabbage/viewitem?obj=1965', so the ID for Cabbage is 1965.";
  var textOutdated = "Important: Sometimes, items will not update. This is normal, and is caused by Jagex's API being unreliable. In the worst case, some item prices might be a few hours outdated.";
  textOutdated += "\nIf you're first setting this up, give it about a day to ensure all items are updated, or manually insert the prices into the sheet.";
  textOutdated += "\nThe 'Last succesful update' and 'Last attempted update' columns can give you an idea of how outdated an item's price is.";
  var ui = SpreadsheetApp.getUi();
  ui.alert("RuneScape Price Updater - Manual", textDescription + "\n\n" + textHowToRefer + "\n\n" + textItemID + "\n\n" + textOutdated, ui.ButtonSet.OK);
}

//Prompt displayed when updating a specific row
function showAboutPrompt_()
{
  var textContact = "This script was written by Zenyl. If you have any questions or feedback, feel free to send me a message on Reddit (/u/zenyl) or in-game (RSN Zenyl).";
  var ui = SpreadsheetApp.getUi();
  ui.alert("RuneScape Price Updater - About", textContact, ui.ButtonSet.OK);
}

function generatePriceUpdateSheet()
{
  //Throw error if not called from a spreadsheet
  if(SpreadsheetApp.getActive() == null)
    throw new Error("This function must be run in relation to a Google Spreadsheet.");
  
  //Throw error if a sheet with {sheetName} already exists
  if(SpreadsheetApp.getActive().getSheetByName(priceSheetName) != null)
    throw new Error("A sheet by the name '" + priceSheetName + "' already exists. Either delete the existing sheet or change the 'priceSheetName' variable at the top of the script.");
  
  //Column headers
  var titles = ["Icon", "Item ID", "Item name", "Price", "Last succesful update", "Last attempted update"];
  
  //Create sheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.insertSheet(priceSheetName);
  
  //Format sheet
  sheet.deleteColumns(1, 25);
  
  for(row = 1; row <= titles.length; row++)
  {
    sheet.getRange(1, row).setValue(titles[row - 1]);
    sheet.getRange(1, row).setFontWeight("bold");
    sheet.getRange(1, row).setHorizontalAlignment("center");
  }
  
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 33);
  sheet.setColumnWidth(3, 150);
  sheet.autoResizeColumn(5);
  sheet.autoResizeColumn(6);
  sheet.deleteRows(499, 501);
  
  //Fill data rows
  for(item = 2; item <= 499; item++)
  {
    sheet.setRowHeight(item, 33);
    sheet.getRange(item, 1).setValue("=IMAGE(\"http://services.runescape.com/m=itemdb_rs/obj_sprite.gif?id=\" & B" + item + ")");
    sheet.getRange(item, 4).setNumberFormat("#,##0 gp")
  }
  
  sheet.setColumnWidth(4, 107);
  
  //Set up triggers
  for(batch = 1; batch <= 10; batch++)
  {
    //Removes existing trigger
    var triggers = ScriptApp.getProjectTriggers();
    for(trigger = 0; trigger < triggers.length; trigger++)
    {
      if(triggers[trigger].getHandlerFunction() == "updateRowsBatch" + batch)
      {
        ScriptApp.deleteTrigger(triggers[trigger]);
      }
    }
    
    //Creates triggers
    ScriptApp.newTrigger("updateRowsBatch" + batch).timeBased().everyHours(1).create();
  }
}

//Trigger functions
function updateRowsBatch1()  { updatePricesForRange_(  2,  49); }
function updateRowsBatch2()  { updatePricesForRange_( 50,  99); }
function updateRowsBatch3()  { updatePricesForRange_(100, 149); }
function updateRowsBatch4()  { updatePricesForRange_(150, 199); }
function updateRowsBatch5()  { updatePricesForRange_(200, 249); }
function updateRowsBatch6()  { updatePricesForRange_(250, 299); }
function updateRowsBatch7()  { updatePricesForRange_(300, 349); }
function updateRowsBatch8()  { updatePricesForRange_(350, 399); }
function updateRowsBatch9()  { updatePricesForRange_(400, 449); }
function updateRowsBatch10() { updatePricesForRange_(450, 499); }

//Updates the item prices from row {from} to row {to}.
function updatePricesForRange_(from, to)
{  
  for(row = from; row <= to; row++)
  {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(priceSheetName);
    var itemID = sheet.getRange(row, 2).getValue();
    if(itemID != "" && isNumber_(itemID) && itemID >= 0)
    {
      Utilities.sleep(5000);
      var newPrice = getPrice_(itemID);
      var sheet = SpreadsheetApp.getActive().getSheetByName(priceSheetName);
      var now = new Date();
      if(newPrice != -1)
      {
        sheet.getRange(row, 4).setValue(newPrice);
        sheet.getRange(row, 5).setValue(now);
      }
      sheet.getRange(row, 6).setValue(now);
    }
  }
}

//Attempts to get the price of an item based on its item ID. Returns -1 if an error occurs.
function getPrice_(id)
{
  try
  {
    var api = "http://services.runescape.com/m=itemdb_rs/api/graph/";
    var i = JSON.parse(UrlFetchApp.fetch(api + id + ".json"))["daily"];
    var e = Object.keys(i).sort().reverse()[0];
    return i[e];
  }
  catch(err)
  {
    return -1;
  }
}

//Returns true if the input is a valid number. Returns false if the input is not a valid number.
function isNumber_(input)
{
  if(parseInt(input).toFixed() == "NaN")
    return false;
  return true;
}

var version = "1.0";
