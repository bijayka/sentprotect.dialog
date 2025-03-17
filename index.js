/*
* Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
* See LICENSE in the project root for license information.
*/
Office.onReady((info) => {
  // Your code that uses Office.js APIs goes here
  console.log("Office.js is ready!");

  function onMessageSendHandler(event) {
    Office.context.mailbox.item.to.getAsync({ asyncContext: event }, getRecipientsCallback);
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
    const externalRecipients = recipients.filter(recipient => {
      const email = recipient.emailAddress.toLowerCase();
      return !email.endsWith("@ey.com") && !email.endsWith("@ey.net");
    });

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
      const message = "Failed to get attachments";
      console.error(message);
      event.completed({ allowEvent: false, errorMessage: message });
      return;
    }

    const attachments = asyncResult.value;
    const nonImageAttachments = attachments.filter(attachment => {
      const extension = attachment.name.split('.').pop().toLowerCase();
      const imageExtensions = ["gif", "jpg", "png", "webp", "tif", "tiff", "jpeg", "jif", "jfif", "jp2", "jpx", "j2k", "j2c"];
      return !imageExtensions.includes(extension);
    });

    if (nonImageAttachments.length > 0) {
      const externalEmails = externalRecipients.map(recipient => recipient.emailAddress).join(", ");
      const attachmentNames = nonImageAttachments.map(attachment => attachment.name).join(", ");
      const message = `
        <p><strong>Warning:</strong> The following recipients are not from the "ey.com" or "ey.net" domains:</p>
        <p>${externalEmails}</p>
        <p>The following attachments are included:</p>
        <p>${attachmentNames}</p>
        <p>Do you want to proceed?</p>
      `;
      event.completed({ allowEvent: false, errorMessage: message });
    } else {
      event.completed({ allowEvent: true });
    }
  }

  // IMPORTANT: To ensure your add-in is supported in Outlook, remember to map the event handler name specified in the manifest to its JavaScript counterpart.
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
});
