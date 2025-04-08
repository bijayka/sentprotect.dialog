// /*
// * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
// * See LICENSE in the project root for license information.
// */
// w_indexjs_globa_var = "Hello World!";
// var extRecipients = [];
// var extAttachments = [];
// let item;

// // Confirms that the Office.js library is loaded.
// Office.onReady((info) => {
//   if (info.host === Office.HostType.Outlook) {
//     console.log("Office.js is ready.");
//   }
// });

// function onMessageSendHandler(event) {
//   // c1
//   event.completed({ allowEvent: false });
//   Office.context.ui.displayDialogAsync(
//     "https://gray-moss-0578a810f.6.azurestaticapps.net/dialog.html",
//     { height: 50, width: 50, displayInIframe: true},
//     (asyncResult) => {
//       if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
//         const dialog = asyncResult.value;

//         // Set a timeout to handle long-running dialogs
//         const timeout = setTimeout(() => {
//           dialog.close();
//           event.completed({ allowEvent: false, errorMessage: "Dialog timed out. Please try again." });
//         }, 30000); // 30 seconds timeout

//         // Handle messages from the dialog
//         dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
//           // clearTimeout(timeout); // Clear timeout on successful message
//           // if (message.message === "allowSend") {
//           //   dialog.close();
//           //   event.completed({ allowEvent: true });
//           // } else if (message.message === "cancelSend") {
//           //   dialog.close();
//           //   event.completed({ allowEvent: false, errorMessage: "Email sending canceled by user." });
//           // }
//           let message;
//                 try {
//                     message = JSON.parse(arg.message);
//                 } catch (e) {
//                     console.error('Error parsing message:', e);
//                     return;
//                 }

//                 if (message.action === "allowSend") {
//                     dialog.close();
//                     // Allow the email to be sent
//                     event.completed({ allowEvent: true });
//                 } else if (message.action === "cancelSend") {
//                     dialog.close();
//                     event.completed({ 
//                         allowEvent: false, 
//                         errorMessage: "Send canceled by user" 
//                     });
//                 }
//         });

//         // Handle dialog closed
//         dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
//           if (arg.error === 12006) {
//             // Dialog was closed
//             event.completed({ 
//               allowEvent: false, 
//               errorMessage: "Dialog was closed" 
//             });
//           }
//           // clearTimeout(timeout); // Clear timeout if dialog is closed
//           // event.completed({ allowEvent: false, errorMessage: "Dialog was closed before confirmation." });
//         });
//       } else {
//         console.error("Failed to open dialog:", asyncResult.error.message);
//         event.completed({ allowEvent: false, errorMessage: "Failed to open confirmation dialog." });
//       }
//     }
//   );
// }

// function getToRecipientsCallback(asyncResult) {
//   const event = asyncResult.asyncContext;
//   if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
//     const message = "Failed to get to recipients";
//     console.error(message);
//     event.completed({ allowEvent: false, errorMessage: message });
//     return;
//   }

//   const recipients = asyncResult.value;
//   externalRecipients = recipients.filter(recipient => {
//     const email = recipient.emailAddress.toLowerCase();
//     return !email.endsWith("@ey.com") && !email.endsWith("@ey.net");
//   });

//   if (externalRecipients.length > 0) {
//     externalRecipients.forEach((recipient, index) => {
//       console.log(`External recipient ${index + 1}: ${recipient.emailAddress}`); 
//       extRecipients.push(recipient.emailAddress); 

//     });
//   } 
//   item.cc.getAsync({ asyncContext: event }, getCCRecipientsCallback);
// }

// function getCCRecipientsCallback(asyncResult) {
//   const event = asyncResult.asyncContext;
//   if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
//     const message = "Failed to get CC recipients";
//     console.error(message);
//     event.completed({ allowEvent: false, errorMessage: message });
//     return;
//   }

//   const recipients = asyncResult.value;
//   externalRecipients = recipients.filter(recipient => {
//     const email = recipient.emailAddress.toLowerCase();
//     return !email.endsWith("@ey.com") && !email.endsWith("@ey.net");
//   });

//   if (externalRecipients.length > 0) {
//     externalRecipients.forEach((recipient, index) => {
//       console.log(`External recipient ${index + 1}: ${recipient.emailAddress}`); 
//       extRecipients.push(recipient.emailAddress); 

//     });
//   } 
//   item.bcc.getAsync({ asyncContext: event }, getBccRecipientsCallback);
// }

// function getBccRecipientsCallback(asyncResult) {
//   const event = asyncResult.asyncContext;
//   if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
//     const message = "Failed to get BCC recipients";
//     console.error(message);
//     event.completed({ allowEvent: false, errorMessage: message });
//     return;
//   }

//   const recipients = asyncResult.value;
//   externalRecipients = recipients.filter(recipient => {
//     const email = recipient.emailAddress.toLowerCase();
//     return !email.endsWith("@ey.com") && !email.endsWith("@ey.net");
//   });

//   if (externalRecipients.length > 0) {
//     externalRecipients.forEach((recipient, index) => {
//       console.log(`External recipient ${index + 1}: ${recipient.emailAddress}`); 
//       extRecipients.push(recipient.emailAddress); 

//     });
//   } 
//   if (extRecipients.length > 0) {
//     item.getAttachmentsAsync({ asyncContext: event }, getAttachmentsCallback2);
//   } else {
//     event.completed({ allowEvent: true });
//   }
// }

// function getAttachmentsCallback2(asyncResult) {
//   const event = asyncResult.asyncContext;
//   if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
//     const message = "Failed to retrieve attachments. Please try again or contact support.";
//     console.error(message);
//     event.completed({ allowEvent: false, errorMessage: message });
//     return;
//   }

//   const attachments = asyncResult.value;

//   if (!attachments || attachments.length === 0) {
//     event.completed({ allowEvent: true });
//     return;
//   }

//   const nonImageAttachments = attachments.filter(attachment => {
//     if (!attachment.name) return false; // Skip attachments without a name
//     const extension = attachment.name.split('.').pop().toLowerCase();
//     const imageExtensions = ["gif", "jpg", "png", "webp", "tif", "tiff", "jpeg", "jif", "jfif", "jp2", "jpx", "j2k", "j2c"];
//     return !imageExtensions.includes(extension);
//   });

//   if (nonImageAttachments.length > 0) {
//     nonImageAttachments.forEach((mAttachmnt, index) => {
//       console.log(`Attachment ${index + 1}: ${mAttachmnt.name}`); 
//       extAttachments.push(mAttachmnt.name); 

//     });

//     let strRecipients = JSON.stringify(extRecipients);
//     localStorage.setItem("strRecipients", strRecipients);
//     let strAttachments = JSON.stringify(extAttachments);
//     localStorage.setItem("strAttachments", strAttachments);

//     event.completed({ allowEvent: false, errorMessage: "Your email includes external recipients with attachment; please review it before sending.", commandId: "msgComposeOpenPaneButton" });
//   } else {
//     event.completed({ allowEvent: true });
//   }
// }

// function getRecipientsCallback(asyncResult) {
//   const event = asyncResult.asyncContext;
//   if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
//     const message = "Failed to get recipients";
//     console.error(message);
//     event.completed({ allowEvent: false, errorMessage: message });
//     return;
//   }

//   const recipients = asyncResult.value;
//   externalRecipients = recipients.filter(recipient => {
//     const email = recipient.emailAddress.toLowerCase();
//     return !email.endsWith("@ey.com") && !email.endsWith("@ey.net");
//   });
//   if (externalRecipients.length > 0) {
//     externalRecipients.forEach((recipient, index) => {
//       console.log(`External recipient ${index + 1}: ${recipient.emailAddress}`);  
//     });
//     Office.context.mailbox.item.getAttachmentsAsync(
//       { asyncContext: { event, externalRecipients } },
//       getAttachmentsCallback);

//   } else {
//     event.completed({ allowEvent: true });
//   }
// }

// function getAttachmentsCallback(asyncResult) {
//   const { event, externalRecipients } = asyncResult.asyncContext;
//   if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
//     const message = "Failed to retrieve attachments. Please try again or contact support.";
//     console.error(message);
//     event.completed({ allowEvent: false, errorMessage: message });
//     return;
//   }

//   const attachments = asyncResult.value;

//   if (!attachments || attachments.length === 0) {
//     event.completed({ allowEvent: true });
//     return;
//   }

//   const nonImageAttachments = attachments.filter(attachment => {
//     if (!attachment.name) return false; // Skip attachments without a name
//     const extension = attachment.name.split('.').pop().toLowerCase();
//     const imageExtensions = ["gif", "jpg", "png", "webp", "tif", "tiff", "jpeg", "jif", "jfif", "jp2", "jpx", "j2k", "j2c"];
//     return !imageExtensions.includes(extension);
//   });

//   if (nonImageAttachments.length > 0) {
//     const externalEmails = externalRecipients.map(recipient => recipient.emailAddress).join("\n- ");
//     const attachmentNames = nonImageAttachments.map(attachment => attachment.name).join("\n- ");
//     const message = `## External Recipients
// A list of external email addresses with checkboxes:
// - ${externalEmails}

// ## Attachments
// A list of file attachments with checkboxes:
// - ${attachmentNames}
//     `;
//     //event.completed({ allowEvent: false, errorMessage: message, sendModeOverride: Office.MailboxEnums.SendModeOverride.PromptUser });
//     event.completed({ allowEvent: false, errorMessage: "Your email includes external recipients with attachment; please review it before sending.", commandId: "msgComposeOpenPaneButton" });
//   } else {
//     event.completed({ allowEvent: true });
//   }
// }

// function handleSendFailure(event, errorMessage) {
//   console.error(errorMessage);
//   event.completed({ allowEvent: false, errorMessage: errorMessage });
// }

// // IMPORTANT: To ensure your add-in is supported in Outlook, remember to map the event handler name specified in the manifest to its JavaScript counterpart.
// Office.actions.associate("onMessageSendHandler", onMessageSendHandler);

/*
* Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
* See LICENSE in the project root for license information.
*/

let extRecipients = [];
let extAttachments = [];
let dialogClosed = false;

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    console.log("Office.js is ready.");
  }
});

function onMessageSendHandler(event) {
  // Prevent immediate sending
  event.completed({ allowEvent: false });

  // Check recipients and attachments first
  checkEmailContent(event).then(() => {
    if (extRecipients.length > 0 && extAttachments.length > 0) {
      // Open dialog for review if there are external recipients and attachments
      openReviewDialog(event);
    } else {
      // Allow sending if no external recipients or attachments
      event.completed({ allowEvent: true });
    }
  }).catch(error => {
    console.error("Error in onMessageSendHandler:", error);
    event.completed({ 
      allowEvent: false, 
      errorMessage: "An error occurred while checking email content." 
    });
  });
}

async function checkEmailContent(event) {
  try {
    // Get all recipients
    const [toRecipients, ccRecipients, bccRecipients] = await Promise.all([
      getRecipientsAsync("to"),
      getRecipientsAsync("cc"),
      getRecipientsAsync("bcc")
    ]);

    // Process recipients
    const allRecipients = [...toRecipients, ...ccRecipients, ...bccRecipients];
    extRecipients = filterExternalRecipients(allRecipients);

    if (extRecipients.length > 0) {
      // Get attachments only if there are external recipients
      const attachments = await getAttachmentsAsync();
      extAttachments = filterNonImageAttachments(attachments);

      // Store in localStorage if needed
      if (extAttachments.length > 0) {
        localStorage.setItem("strRecipients", JSON.stringify(extRecipients.map(r => r.emailAddress)));
        localStorage.setItem("strAttachments", JSON.stringify(extAttachments.map(a => a.name)));
      }
    }
  } catch (error) {
    throw new Error(`Failed to check email content: ${error.message}`);
  }
}

function getRecipientsAsync(field) {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item[field].getAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(new Error(`Failed to get ${field} recipients`));
      }
    });
  });
}

function getAttachmentsAsync() {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.getAttachmentsAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value || []);
      } else {
        reject(new Error("Failed to get attachments"));
      }
    });
  });
}

function filterExternalRecipients(recipients) {
  return recipients.filter(recipient => {
    const email = recipient.emailAddress.toLowerCase();
    return !email.endsWith("@ey.com") && !email.endsWith("@ey.net");
  });
}

function filterNonImageAttachments(attachments) {
  const imageExtensions = ["gif", "jpg", "png", "webp", "tif", "tiff", "jpeg", "jif", "jfif", "jp2", "jpx", "j2k", "j2c"];
  return attachments.filter(attachment => {
    if (!attachment.name) return false;
    const extension = attachment.name.split('.').pop().toLowerCase();
    return !imageExtensions.includes(extension);
  });
}

function openReviewDialog(event) {
  Office.context.ui.displayDialogAsync(
    "https://gray-moss-0578a810f.6.azurestaticapps.net/dialog.html",
    { height: 50, width: 50, displayInIframe: true },
    (asyncResult) => {
      if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        console.error("Failed to open dialog:", asyncResult.error.message);
        event.completed({ 
          allowEvent: false, 
          errorMessage: "Failed to open confirmation dialog" 
        });
        return;
      }

      const dialog = asyncResult.value;
      
      // Set timeout for dialog
      const timeout = setTimeout(() => {
        if (!dialogClosed) {
          dialog.close();
          event.completed({ 
            allowEvent: false, 
            errorMessage: "Dialog timed out. Please try again." 
          });
        }
      }, 30000);

      // Handle messages from dialog
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
        try {
          const message = JSON.parse(arg.message);
          handleDialogMessage(message, dialog, event);
          clearTimeout(timeout);
        } catch (error) {
          console.error("Error processing dialog message:", error);
        }
      });

      // Handle dialog closed
      dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
        if (arg.error === 12006 && !dialogClosed) {
          dialogClosed = true;
          clearTimeout(timeout);
          event.completed({ 
            allowEvent: false, 
            errorMessage: "Dialog was closed" 
          });
        }
      });
    }
  );
}

function handleDialogMessage(message, dialog, event) {
  if (dialogClosed) return;
  dialogClosed = true;

  switch (message.action) {
    case "allowSend":
      dialog.close();
      event.completed({ allowEvent: true });
      break;
    case "cancelSend":
      dialog.close();
      event.completed({ 
        allowEvent: false, 
        errorMessage: "Send canceled by user" 
      });
      break;
    default:
      console.warn("Unknown message action:", message.action);
      break;
  }
}
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);