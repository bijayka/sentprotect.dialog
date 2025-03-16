Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
      // Add the event handler for ItemSend
      Office.context.mailbox.addHandlerAsync(Office.EventType.ItemSend, itemSendHandler);
  }
});

function itemSendHandler(eventArgs) {
  const item = eventArgs.mailboxItem;
  const attachments = item.attachments;
  const recipientDomains = item.to.map(email => email.split('@')[1]);

  // Check if the recipient is external
  const isExternal = recipientDomains.some(domain => !isInternalDomain(domain));

  if (isExternal && attachments.length > 0) {
      eventArgs.completed({ allowEvent: false }); // Prevent sending
      showAlert(attachments, recipientDomains);
  } else {
      eventArgs.completed({ allowEvent: true }); // Allow sending
  }
}

function isInternalDomain(domain) {
  // Define your internal domains here
  const internalDomains = ['yourcompany.com'];
  return internalDomains.includes(domain);
}

function showAlert(attachments, domains) {
  const attachmentList = attachments.map(att => att.name).join(', ');
  const domainList = domains.join(', ');

  // Create a modal dialog to show the alert
  const alertHtml = `
      <div id="alertModal" style="padding: 20px; background: white; border: 1px solid #ccc; border-radius: 5px; width: 300px;">
          <h3>Warning: External Email</h3>
          <p>You are about to send an email with the following attachments:</p>
          <p><strong>${attachmentList}</strong></p>
          <p>To external domains: <strong>${domainList}</strong></p>
          <button id="deferButton">Defer Sending</button>
          <button id="sendButton">Send Anyway</button>
      </div>
  `;

  document.body.insertAdjacentHTML('beforeend', alertHtml);

  // Add event listeners for buttons
  document.getElementById('deferButton').onclick = () => {
      document.getElementById('alertModal').remove();
      // Logic to defer sending (e.g., save draft)
  };

  document.getElementById('sendButton').onclick = () => {
      document.getElementById('alertModal').remove();
      // Allow sending the email
      eventArgs.completed({ allowEvent: true });
  };
}