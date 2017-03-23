Office.initialize = function () {
}

// Helper function to add a status message to the info bar.
function statusUpdate(icon, text) {
  Office.context.mailbox.item.notificationMessages.replaceAsync("status", {
    type: "informationalMessage",
    icon: icon,
    message: text,
    persistent: false
  });
}

function addTextToBody(text, icon, event) {
    Office.context.mailbox.item.body.setSelectedDataAsync(text, { coercionType: Office.CoercionType.Text },
      function (asyncResult) {
          if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
              statusUpdate(icon, "\"" + text + "\" inserted successfully.");
          }
          else {
              Office.context.mailbox.item.notificationMessages.addAsync("addTextError", {
                  type: "errorMessage",
                  message: "Failed to insert \"" + text + "\": " + asyncResult.error.message
              });
          }
          event.completed();
      });
}

// Gets the subject of the item and displays it in the info bar.
function getSubject(event) {
    var subject = Office.context.mailbox.item.subject;

    Office.context.mailbox.item.notificationMessages.addAsync("subject", {
        type: "informationalMessage",
        icon: "blue-icon-16",
        message: "Subject: " + subject,
        persistent: false
    });

    event.completed();
}