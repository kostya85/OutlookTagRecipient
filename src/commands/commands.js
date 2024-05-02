/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

Office.initialize = function (reason) {};

let recipients = '';
let notificationCreated = false;

/**
 * Handles the OnMessageRecipientsChanged event.
 * @param {*} event The Office event object
 */
function tagExternal_onMessageRecipientsChangedHandler(event) {
  console.log("tagExternal_onMessageRecipientsChangedHandler method"); //debugging
  console.log("event: " + JSON.stringify(event)); //debugging
  if (event.changedRecipientFields.to) {
    checkForExternalTo();
  }
  if (event.changedRecipientFields.cc) {
    checkForExternalCc();
  }
  if (event.changedRecipientFields.bcc) {
    checkForExternalBcc();
  }
}

/**
 * Determines if there are any external recipients in the To field.
 */
function checkForExternalTo() {
  console.log("checkForExternalTo method"); //debugging

  // Get To recipients.
  console.log("Get To recipients"); //debugging
  Office.context.mailbox.item.to.getAsync(
    function (asyncResult) {
      // Handle success or error.
      if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        console.error("Failed to get To recipients. " + JSON.stringify(asyncResult.error));
        return;
      }

      const toRecipients = JSON.stringify(asyncResult.value);

      recipients = asyncResult.value;

      console.log("To recipients: " + toRecipients); //debugging
      const keyName = "tagExternalTo";
      if (toRecipients != null
          && toRecipients.length > 0
          && JSON.stringify(toRecipients).includes(Office.MailboxEnums.RecipientType.ExternalUser)) {
        console.log("To includes external users"); //debugging
        _setSessionData(keyName, true);
      } else {
        _setSessionData(keyName, false);
      }
    });
}
/**
 * Determines if there are any external recipients in the Cc field.
 */
function checkForExternalCc() {
  console.log("checkForExternalCc method"); //debugging

  // Get Cc recipients.
  console.log("Get Cc recipients"); //debugging
  Office.context.mailbox.item.cc.getAsync(
    function (asyncResult) {
      // Handle success or error.
      if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        console.error("Failed to get Cc recipients. " + JSON.stringify(asyncResult.error));
        return;
      }
      
      const ccRecipients = JSON.stringify(asyncResult.value);
      console.log("Cc recipients: " + ccRecipients); //debugging
      const keyName = "tagExternalCc";
      if (ccRecipients != null
          && ccRecipients.length > 0
          && JSON.stringify(ccRecipients).includes(Office.MailboxEnums.RecipientType.ExternalUser)) {
        console.log("Cc includes external users"); //debugging
        _setSessionData(keyName, true);
      } else {
        _setSessionData(keyName, false);
      }
    });
}
/**
 * Determines if there are any external recipients in the Bcc field.
 */
function checkForExternalBcc() {
  console.log("checkForExternalBcc method"); //debugging

  // Get Bcc recipients.
  console.log("Get Bcc recipients"); //debugging
  Office.context.mailbox.item.bcc.getAsync(
    function (asyncResult) {
      // Handle success or error.
      if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        console.error("Failed to get Bcc recipients. " + JSON.stringify(asyncResult.error));
        return;
      }

      const bccRecipients = JSON.stringify(asyncResult.value);
      console.log("Bcc recipients: " + bccRecipients); //debugging
      const keyName = "tagExternalBcc";
      if (bccRecipients != null
          && bccRecipients.length > 0
          && JSON.stringify(bccRecipients).includes(Office.MailboxEnums.RecipientType.ExternalUser)) {
        console.log("Bcc includes external users"); //debugging
        _setSessionData(keyName, true);
      } else {
        _setSessionData(keyName, false);
      }
    });
}
/**
 * Sets the value of the specified sessionData key.
 * If value is true, also tag as external, else check entire sessionData property bag.
 * @param {string} key The key or name
 * @param {bool} value The value to assign to the key
 */
 function _setSessionData(key, value) {
  Office.context.mailbox.item.sessionData.setAsync(
    key,
    value.toString(),
    function(asyncResult) {
      // Handle success or error.
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      console.log(`sessionData.setAsync(${key}) to ${value} succeeded`);
      if (value) {
        _tagExternal(value);
      } else {
        _checkForExternal();
      }
    } else {
      console.error(`Failed to set ${key} sessionData to ${value}. Error: ${JSON.stringify(asyncResult.error)}`);
      return;
    }
  });
}
/**
 * Checks the sessionData property bag to determine if any field contains external recipients.
 */
function _checkForExternal() {
  console.log("_checkForExternal method"); //debugging

  // Get sessionData to determine if any fields have external recipients.
  Office.context.mailbox.item.sessionData.getAllAsync(
    function (asyncResult) {
      // Handle success or error.
      if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        console.error("Failed to get all sessionData. " + JSON.stringify(asyncResult.error));
        return;
      }

      const sessionData = JSON.stringify(asyncResult.value);
      console.log("Current SessionData: " + sessionData); //debugging
      if (sessionData != null
        && sessionData.length > 0
        && sessionData.includes("true")) {
        console.log("At least one recipients field includes external users"); //debugging
        _tagExternal(true);
      } else {
        console.log("There are no external recipients"); //debugging
        _tagExternal(false);
      }
  });
}
/**
 * If there are any external recipients, prepends the subject of the Outlook item
 * with "[External]" and appends a disclaimer to the item body. If there are
 * no external recipients, ensures the tag is not present and clears the disclaimer.
 * @param {bool} hasExternal If there are any external recipients
 */
function _tagExternal(hasExternal) {
  console.log("_tagExternal method"); //debugging

  // External subject tag.
  let externalTag = "[External]";

  if (hasExternal) {
    try {
      externalTag += JSON.stringify(Office.context.mailbox.item.notificationMessages);
      externalTag += typeof Office.context.mailbox.item.notificationMessages;
      externalTag += typeof Office.context.mailbox.item.notificationMessages.addAsync;

      let message = '';
      message += 'В списке отправителей обнаружены внешние почтовые адреса';

      if(recipients.length){
        message += ': ';
        for(let i = 0; i < recipients.length; i++){
          if(
            (recipients[i]['recipientType'] !== undefined) &&
            (recipients[i]['recipientType'] === Office.MailboxEnums.RecipientType.ExternalUser)
          ){
            message += recipients[i]['emailAddress'];
          }
        }
      }

      const id = 'kbnotification';
      if(notificationCreated){
        Office.context.mailbox.item.notificationMessages.removeAsync(id, () => {
          notificationCreated = false;
          const details =
          {
              type: Office.MailboxEnums.ItemNotificationMessageType.ProgressIndicator,
              message: message.slice(0, 145) + '...'
          };
          Office.context.mailbox.item.notificationMessages.addAsync(id, details, () => {
            notificationCreated = true;
          });
          });
      }else{
        const details =
        {
            type: Office.MailboxEnums.ItemNotificationMessageType.ProgressIndicator,
            message: message.slice(0, 145) + '...'
        };
        Office.context.mailbox.item.notificationMessages.addAsync(id, details, () => {
          notificationCreated = true;
        });
      }

    } catch (err) {
      externalTag += 'has error';
      externalTag += err.message;
      externalTag += JSON.stringify(err);
    }

    console.log("External: Get Subject"); //debugging
    
    // Ensure "[External]" is prepended to the subject.
    Office.context.mailbox.item.subject.getAsync(
      function (asyncResult) {
        // Handle success or error.
        if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
          console.error("Failed to get subject. " + JSON.stringify(asyncResult.error));
          return;
        }

        console.log("Current Subject: " + JSON.stringify(asyncResult.value)); //debugging
        let subject = asyncResult.value;
        if (!subject.includes(externalTag)) {
          subject = `${externalTag} ${subject}`;
          console.log("Updated Subject: " + subject); //debugging
          Office.context.mailbox.item.subject.setAsync(
            subject,
            function (asyncResult) {
              // Handle success or error.
              if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
                console.error("Failed to set Subject. " + JSON.stringify(asyncResult.error));
                return;
              }

              console.log("Set subject succeeded"); //debugging
          });
        }
    });
  } else {
    try {
      if(notificationCreated){
        const id = 'kbnotification';
        Office.context.mailbox.item.notificationMessages.removeAsync(id, () => {
          notificationCreated = false;
        });
      }
    } catch (err) {
    }
  }
}

// 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
Office.actions.associate("tagExternal_onMessageRecipientsChangedHandler", tagExternal_onMessageRecipientsChangedHandler);