/*
* Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
* See LICENSE in the project root for license information.
*/
w_indexjs_globa_var = "Hello World!";
var extRecipients = [];
var extAttachments = [];
let item;

// Confirms that the Office.js library is loaded.
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
      item = Office.context.mailbox.item;
      console.log('Office is ready');
  }
});

// Gets the email addresses of all the recipients of the item being composed.
function getAllRecipients() {
  let toRecipients, ccRecipients, bccRecipients;


  // Verify if the mail item is an appointment or message.
  if (item.itemType === Office.MailboxEnums.ItemType.Appointment) {
      toRecipients = item.requiredAttendees;
      ccRecipients = item.optionalAttendees;
      console.log('item.itemType');
      console.log(item.itemType);
  }
  else {
      toRecipients = item.to;
      ccRecipients = item.cc;
      bccRecipients = item.bcc;
  }

  // Get the recipients from the To or Required field of the item being composed.
  toRecipients.getAsync((asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.log(asyncResult.error.message);
          return;
      }
      addAddresses(asyncResult.value);
  });

  // Get the recipients from the Cc or Optional field of the item being composed.
  ccRecipients.getAsync((asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.log(asyncResult.error.message);
          return;
      }
      addAddresses(asyncResult.value);

  });

  // Get the recipients from the Bcc field of the message being composed, if applicable.
  // if (bccRecipients.length > 0) {
  //     bccRecipients.getAsync((asyncResult) => {
  //     if (asyncResult.status === Office.AsyncResultStatus.Failed) {
  //         write(asyncResult.error.message);
  //         return;
  //     }
  //     addAddresses(asyncResult.value);

  //     });
  // } else {
  //     console.log("Recipients in the Bcc field: None");
  // }

  item.bcc.getAsync(function(asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      const msgBcc = asyncResult.value;
      for (let i = 0; i < msgBcc.length; i++) {
        console.log(msgBcc[i].displayName + " (" + msgBcc[i].emailAddress + ")");
        extRecipients.push(msgBcc[i].emailAddress);
      }
    } else {
      console.error(asyncResult.error);
      return;
    }
  });
}

function addAddresses (recipients) {
  for (let i = 0; i < recipients.length; i++) {
    extRecipients.push(recipients[i].emailAddress);
  }
}


  function onMessageSendHandler(event) {
    console.warn(w_indexjs_globa_var);
    Office.context.mailbox.item.to.getAsync({ asyncContext: event }, getToRecipientsCallback);
  }

  function myAttachments() {
    const item2 = Office.context.mailbox.item2;

    if (item2.attachments.length > 0) {
        for (let i = 0; i < item2.attachments.length; i++) {
            const attachment = item2.attachments[i];
            console.log(`${i+1}. Name: ${attachment.name}`);
            console.log(`ID: ${attachment.id}`);
            console.log(`Type: ${attachment.attachmentType}`);
            console.log(`Inline content: ${attachment.isInline}`);
            console.log(`Size: ${attachment.size}`);
        }
    } else {
        console.log("This mail item2 doesn't contain any attachments.");
    }
}
  

  function sampleFunction() {
    console.log('sampleFunction');
    Office.context.mailbox.item.getAttachmentsAsync(function(asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        const attachments = asyncResult.value;
        if (attachments.length > 0) {
          for (let i = 0; i < attachments.length; i++) {
            console.log(`${i+1}. Name: ${attachments[i].name}`);
            console.log(`ID: ${attachments[i].id}`);
            console.log(`Type: ${attachments[i].attachmentType}`);
            console.log(`Inline content: ${attachments[i].isInline}`);
            console.log(`Size: ${attachments[i].size}`);
          }
        } else {
          console.log("This mail item doesn't contain any attachments.");
        }
      } else {
        console.error(asyncResult.error);
      }
    });
  }

  function getBcc() {
    Office.context.mailbox.item.bcc.getAsync(function(asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        const msgBcc = asyncResult.value;
        const externalBCC = msgBcc.filter(recipient => {
          const email = recipient.emailAddress.toLowerCase();
          return !email.endsWith("@ey.com") && !email.endsWith("@ey.net");
        });
        console.log("Message being blind-copied to:");
        for (let i = 0; i < externalBCC.length; i++) {
          console.log(externalBCC[i].displayName + " (" + externalBCC[i].emailAddress + ")");
        }
      } else {
        console.error(asyncResult.error);
        const event = asyncResult.asyncContext;
        event.completed({ allowEvent: false, errorMessage: message });
      }
    });
  }

  function getToRecipientsCallback(asyncResult) {
    const event = asyncResult.asyncContext;
    if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
      const message = "Failed to get to recipients";
      console.error(message);
      event.completed({ allowEvent: false, errorMessage: message });
      return;
    }

    const recipients = asyncResult.value;
    externalRecipients = recipients.filter(recipient => {
      const email = recipient.emailAddress.toLowerCase();
      return !email.endsWith("@ey.com") && !email.endsWith("@ey.net");
    });

    if (externalRecipients.length > 0) {
      externalRecipients.forEach((recipient, index) => {
        console.log(`External recipient ${index + 1}: ${recipient.emailAddress}`); 
        extRecipients.push(recipient.emailAddress); 

      });
    } 
    item.cc.getAsync({ asyncContext: event }, getCCRecipientsCallback);
  }

  function getCCRecipientsCallback(asyncResult) {
    const event = asyncResult.asyncContext;
    if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
      const message = "Failed to get CC recipients";
      console.error(message);
      event.completed({ allowEvent: false, errorMessage: message });
      return;
    }

    const recipients = asyncResult.value;
    externalRecipients = recipients.filter(recipient => {
      const email = recipient.emailAddress.toLowerCase();
      return !email.endsWith("@ey.com") && !email.endsWith("@ey.net");
    });

    if (externalRecipients.length > 0) {
      externalRecipients.forEach((recipient, index) => {
        console.log(`External recipient ${index + 1}: ${recipient.emailAddress}`); 
        extRecipients.push(recipient.emailAddress); 

      });
    } 
    item.bcc.getAsync({ asyncContext: event }, getBccRecipientsCallback);
  }

  function getBccRecipientsCallback(asyncResult) {
    const event = asyncResult.asyncContext;
    if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
      const message = "Failed to get BCC recipients";
      console.error(message);
      event.completed({ allowEvent: false, errorMessage: message });
      return;
    }

    const recipients = asyncResult.value;
    externalRecipients = recipients.filter(recipient => {
      const email = recipient.emailAddress.toLowerCase();
      return !email.endsWith("@ey.com") && !email.endsWith("@ey.net");
    });

    if (externalRecipients.length > 0) {
      externalRecipients.forEach((recipient, index) => {
        console.log(`External recipient ${index + 1}: ${recipient.emailAddress}`); 
        extRecipients.push(recipient.emailAddress); 

      });
    } 
    if (extRecipients.length > 0) {
      item.getAttachmentsAsync({ asyncContext: event }, getAttachmentsCallback2);
    } else {
      event.completed({ allowEvent: true });
    }
  }

  function getAttachmentsCallback2(asyncResult) {
    const { event } = asyncResult.asyncContext;
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
      nonImageAttachments.forEach((mAttachmnt, index) => {
        console.log(`Attachment ${index + 1}: ${mAttachmnt.name}`); 
        extAttachments.push(mAttachmnt.name); 

      });
//       const externalEmails = externalRecipients.map(recipient => recipient.emailAddress).join("\n- ");
//       const attachmentNames = nonImageAttachments.map(attachment => attachment.name).join("\n- ");
//       const message = `## External Recipients
// A list of external email addresses with checkboxes:
// - ${externalEmails}

// ## Attachments
// A list of file attachments with checkboxes:
// - ${attachmentNames}
//       `;
      //event.completed({ allowEvent: false, errorMessage: message, sendModeOverride: Office.MailboxEnums.SendModeOverride.PromptUser });
      event.completed({ allowEvent: false, errorMessage: "Your email includes external recipients with attachment; please review it before sending.", commandId: "msgComposeOpenPaneButton" });
    } else {
      event.completed({ allowEvent: true });
    }
  }


  function getRecipientsCallback(asyncResult) {
    const event = asyncResult.asyncContext;
    if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
      const message = "Failed to get recipients";
      console.error(message);
      event.completed({ allowEvent: false, errorMessage: message });
      return;
    }

    const recipients = asyncResult.to.value .value;
    externalRecipients = recipients.filter(recipient => {
      const email = recipient.emailAddress.toLowerCase();
      return !email.endsWith("@ey.com") && !email.endsWith("@ey.net");
    });
    if (externalRecipients.length > 0) {
      externalRecipients.forEach((recipient, index) => {
        console.log(`External recipient ${index + 1}: ${recipient.emailAddress}`);  
      });
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

  function handleSendFailure(event, errorMessage) {
    console.error(errorMessage);
    event.completed({ allowEvent: false, errorMessage: errorMessage });
  }

  // IMPORTANT: To ensure your add-in is supported in Outlook, remember to map the event handler name specified in the manifest to its JavaScript counterpart.
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);

