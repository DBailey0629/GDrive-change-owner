/**
 * To Install: Share with a new user as editor. They must open the script, click
 * "Deploy > Test Deployment > Install"
 */

/**
 * The maximum number of characters that can fit in the cat image.
 */
var MAX_MESSAGE_LENGTH = 40;

/**
 * Callback for rendering the homepage card.
 * @return {CardService.Card} The card to show to the user.
 */
function onHomepage(e) {
  console.log(e);
  var hour = Number(Utilities.formatDate(new Date(), e.userTimezone.id, 'H'));
  var message;
  if (hour >= 6 && hour < 12) {
    message = 'Good morning';
  } else if (hour >= 12 && hour < 18) {
    message = 'Good afternoon';
  } else {
    message = 'Good night';
  }
  message += ' ' + e.hostApp;
  return createCatCard(message, true);
}

/**
 * Creates a card with an image of a cat, overlayed with the text.
 * @param {String} text The text to overlay on the image.
 * @param {Boolean} isHomepage True if the card created here is a homepage;
 *      false otherwise. Defaults to false.
 * @return {CardService.Card} The assembled card.
 */
function createCatCard(text, isHomepage) {
  // Explicitly set the value of isHomepage as false if null or undefined.
  if (!isHomepage) {
    isHomepage = false;
  }

  // Use the "Cat as a service" API to get the cat image. Add a "time" URL
  // parameter to act as a cache buster.
  var now = new Date();
  // Replace formward slashes in the text, as they break the CataaS API.
  var caption = text.replace(/\//g, ' ');
  var imageUrl =
      Utilities.formatString('https://cataas.com/cat/says/%s?time=%s',
          encodeURIComponent(caption), now.getTime());
  var image = CardService.newImage()
      .setImageUrl(imageUrl)
      .setAltText('Meow')

  // Create a button that changes the cat image when pressed.
  // Note: Action parameter keys and values must be strings.
  var action = CardService.newAction()
      .setFunctionName('onChangeCat')
      .setParameters({text: text, isHomepage: isHomepage.toString()});
  var button = CardService.newTextButton()
      .setText('Change cat')
      .setOnClickAction(action)
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED);
  var buttonSet = CardService.newButtonSet()
      .addButton(button);

  // Create a footer to be shown at the bottom.
  var footer = CardService.newFixedFooter()
      .setPrimaryButton(CardService.newTextButton()
          .setText('Powered by cataas.com')
          .setOpenLink(CardService.newOpenLink()
              .setUrl('https://cataas.com')));

  // Assemble the widgets and return the card.
  var section = CardService.newCardSection()
      .addWidget(image)
      .addWidget(buttonSet);
  var card = CardService.newCardBuilder()
      .addSection(section)
      .setFixedFooter(footer);

  if (!isHomepage) {
    // Create the header shown when the card is minimized,
    // but only when this card is a contextual card. Peek headers
    // are never used by non-contexual cards like homepages.
    var peekHeader = CardService.newCardHeader()
      .setTitle('Contextual Cat')
      .setImageUrl('https://www.gstatic.com/images/icons/material/system/1x/pets_black_48dp.png')
      .setSubtitle(text);
    card.setPeekCardHeader(peekHeader)
  }

  return card.build();
}

/**
 * Callback for the "Change cat" button.
 * @param {Object} e The event object, documented {@link
 *     https://developers.google.com/gmail/add-ons/concepts/actions#action_event_objects
 *     here}.
 * @return {CardService.ActionResponse} The action response to apply.
 */
function onChangeCat(e) {
  console.log(e);
  // Get the text that was shown in the current cat image. This was passed as a
  // parameter on the Action set for the button.
  var text = e.parameters.text;

  // The isHomepage parameter is passed as a string, so convert to a Boolean.
  var isHomepage = e.parameters.isHomepage === 'true';

  // Create a new card with the same text.
  var card = createCatCard(text, isHomepage);

  // Create an action response that instructs the add-on to replace
  // the current card with the new one.
  var navigation = CardService.newNavigation()
      .updateCard(card);
  var actionResponse = CardService.newActionResponseBuilder()
      .setNavigation(navigation);
  return actionResponse.build();
}

/**
 * Truncate a message to fit in the cat image.
 * @param {string} message The message to truncate.
 * @return {string} The truncated message.
 */
function truncate(message) {
  if (message.length > MAX_MESSAGE_LENGTH) {
    message = message.slice(0, MAX_MESSAGE_LENGTH);
    message = message.slice(0, message.lastIndexOf(' ')) + '...';
  }
  return message;
}


/**
 * Callback for rendering the card for specific Drive items.
 * @param {Object} e The event object.
 * @return {CardService.Card} The card to show to the user.
 */
function onDriveItemsSelected(e) {
  Logger.log(e);
  var items = e.drive.selectedItems;
  // Include at most 10 items in the text.
  items = items.slice(0, 10);
  var text = items.map(function(item) {
    var title = item.title;
    // If neccessary, truncate the title to fit in the image.
    title = truncate(title);
    Logger.log(item);
    Logger.log(DriveApp.getFileById(item.id).getOwner().getEmail());
    if (true == true ) {
      if (Drive.Comments.list(item.id).items.length>0) {
        let newId = convToMicrosoft(item.id);
        return "This document has Comments.";
      }
      //Drive.Files.copy({'setModifiedDate':true, 'title':item.title, 'parents': [{'id': DriveApp.getFileById(item.id).getParents().next().getId()}], 'modifiedDate': Drive.Files.get(item.id).modifiedDate}, item.id); //works but doesn't copy comments
      /*let newId = Drive.Files.copy({'title':item.title, 'parents': [{'id': DriveApp.getFileById(item.id).getParents().next().getId()}] }, item.id).id; //works but doesn't copy comments
      let url = "https://www.googleapis.com/drive/v3/files/" + newId ;
      let params = {
        method: "patch",
        headers: {"Authorization": "Bearer " + ScriptApp.getOAuthToken()},
        payload: JSON.stringify({modifiedTime: Drive.Files.get(item.id).modifiedDate}),
        contentType: "application/json"
      };
      UrlFetchApp.fetch(url, params);*/

      let url = "https://www.googleapis.com/drive/v3/files/" + item.id + "/copy";
      let params = {
        method: "post",
        headers: {"Authorization": "Bearer " + ScriptApp.getOAuthToken()},
        payload: JSON.stringify({name: item.title}),
        contentType: "application/json"
      };
      let newId = JSON.parse(UrlFetchApp.fetch(url, params)).id;
      Utilities.sleep(50);
      let count = 0;
      while(Drive.Files.get(item.id).modifiedDate !== Drive.Files.get(newId).modifiedDate && count <=20) {
        Logger.log(newId);
        url = "https://www.googleapis.com/drive/v3/files/" + newId ;
        let params2 = {
          method: "patch",
          headers: {"Authorization": "Bearer " + ScriptApp.getOAuthToken()},
          payload: JSON.stringify({modifiedTime: Drive.Files.get(item.id).modifiedDate}),
          contentType: "application/json",
          muteHttpExceptions: true
        };
        UrlFetchApp.fetch(url, params2);
        Utilities.sleep(50);
        count++;
      }
      
      

      
      //Logger.log(Drive.Files.get(newId).properties);

    }
    return title;
  }).join('\n');
  return createCatCard(text);
}

function convToMicrosoft(fileId) {
  if (fileId == null) throw new Error("No file ID.");
  var file = DriveApp.getFileById(fileId);
  var mime = file.getMimeType();
  var format = "";
  var ext = "";
  switch (mime) {
    case "application/vnd.google-apps.document":
      format = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
      ext = ".docx";
      break;
    case "application/vnd.google-apps.spreadsheet":
      format = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
      ext = ".xlsx";
      break;
    case "application/vnd.google-apps.presentation":
      format = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
      ext = ".pptx";
      break;
    default:

      let url = "https://www.googleapis.com/drive/v3/files/" + fileId + "/copy";
      let params = {
        method: "post",
        headers: {"Authorization": "Bearer " + ScriptApp.getOAuthToken()},
        payload: JSON.stringify({modifiedTime: Drive.Files.get(fileId).modifiedDate}),
        contentType: "application/json"
      };
      UrlFetchApp.fetch(url, params);
      return null;
  }
  
  var url = "https://www.googleapis.com/drive/v3/files/" + fileId + "/export?mimeType=" + format;
  var blob = UrlFetchApp.fetch(url, {
    method: "get",
    headers: {"Authorization": "Bearer " + ScriptApp.getOAuthToken()},
    muteHttpExceptions: true
  }).getBlob();
  var filename = file.getName();
  //Logger.log(blob);
  var newId = DriveApp.createFile(blob).setName(~filename.indexOf(ext) ? filename : filename + ext).getId();
  DriveApp.getFileById(newId).moveTo(DriveApp.getFileById(fileId).getParents().next());
  
  url = "https://www.googleapis.com/drive/v3/files/" + newId ;
  let params = {
    method: "patch",
    headers: {"Authorization": "Bearer " + ScriptApp.getOAuthToken()},
    payload: JSON.stringify({modifiedTime: Drive.Files.get(fileId).modifiedDate}),
    contentType: "application/json"
  };
  UrlFetchApp.fetch(url, params);
    

  //Logger.log(JSON.parse(newId).id); 
  /*url = "https://www.googleapis.com/drive/v2/files/" + newId + "?setModifiedDate=true&updateViewedDate=false";
  UrlFetchApp.fetch(url, {
    method: "put",
    
    headers: {"Authorization": "Bearer " + ScriptApp.getOAuthToken()},
    payload: JSON.stringify({"title": "Whos yo daddy", "modifiedDate": }),//Drive.Files.get(fileId).modifiedDate }),
    muteHttpExceptions: false
  });*/
  //var filename = file.getName();
  //var id = DriveApp.createFile(blob).setName(~filename.indexOf(ext) ? filename : filename + ext).getId();
  return newId;
};
