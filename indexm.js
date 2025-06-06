/*
* Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
* See LICENSE in the project root for license information.
*/
// w_indexjs_globa_var = "Hello World!";
var externalRecipients = [];
Office.onReady((info) => {
  // Your code that uses Office.js APIs goes here
  console.log("Office.js is ready!");
  console.log('test');


  //Modifications
  console.log("Opening dialog...");
 

  function onMessageSendHandler(event) {
    Office.context.ui.displayTaskPaneAsync(
      "https://gray-moss-0578a810f.6.azurestaticapps.net/dialog.html", // URL of the task pane
      { title: "Review Email" }, // Task pane title
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
          console.log("Task pane opened successfully.");
          event.completed({ allowEvent: false, errorMessage: "Please review the email in the task pane before sending." });
        } else {
          console.error("Failed to open task pane:", asyncResult.error.message);
          event.completed({ allowEvent: false, errorMessage: "Failed to open the task pane." });
        }
      }
    );
  }

  function getRecipientsCallback(asyncResult) {
    const event = asyncResult.asyncContext;
    if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
      const message = "Failed to get recipients";
      console.error(message);
      event.completed({ allowEvent: false, errorMessage: message });
      return;
    }

    const recipients = asyncResult.value;
    externalRecipients = recipients.filter(recipient => {
      const email = recipient.emailAddress.toLowerCase();
      return !email.endsWith("@ey.com") && !email.endsWith("@ey.net");
    });
    console.log('externalRecipients');
    console.log(externalRecipients);
    if (externalRecipients.length > 0) {

      Office.context.mailbox.item.getAttachmentsAsync(
        { asyncContext: { event, externalRecipients } },
        getAttachmentsCallback);

    } else {
      event.completed({ allowEvent: true });
    }
  }

  function getAttachmentsCallback(asyncResult) {
    const { event, externalRecipients } = asyncResult.asyncContext;
    if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
      const message = "Failed to retrieve attachments. Please try again or contact support.";
      console.error(message);
      event.completed({ allowEvent: false, errorMessage: message });
      return;
    }

    const attachments = asyncResult.value;

    if (!attachments || attachments.length === 0) {
      event.completed({ allowEvent: true });
      return;
    }

    const nonImageAttachments = attachments.filter(attachment => {
      if (!attachment.name) return false; // Skip attachments without a name
      const extension = attachment.name.split('.').pop().toLowerCase();
      const imageExtensions = ["gif", "jpg", "png", "webp", "tif", "tiff", "jpeg", "jif", "jfif", "jp2", "jpx", "j2k", "j2c"];
      return !imageExtensions.includes(extension);
    });

    if (nonImageAttachments.length > 0) {
      const externalEmails = externalRecipients.map(recipient => recipient.emailAddress).join("\n- ");
      const attachmentNames = nonImageAttachments.map(attachment => attachment.name).join("\n- ");
      const message = `## External Recipients
A list of external email addresses with checkboxes:
- ${externalEmails}

## Attachments
A list of file attachments with checkboxes:
- ${attachmentNames}
      `;
      //event.completed({ allowEvent: false, errorMessage: message, sendModeOverride: Office.MailboxEnums.SendModeOverride.PromptUser });
      event.completed({ allowEvent: false, errorMessage: "Your email includes external recipients with attachment; please review it before sending.", commandId: "msgComposeOpenPaneButton" });
    } else {
      event.completed({ allowEvent: true });
    }
  }

  // IMPORTANT: To ensure your add-in is supported in Outlook, remember to map the event handler name specified in the manifest to its JavaScript counterpart.
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
});



