// commands.js — Client Code Reminder
// Fires on every send. If the subject has no [CODE] prefix, cancels the send
// and shows a reminder. If the user clicks Send again without adding a code,
// the bypass flag allows the send through.

Office.initialize = function () {};

function onItemSend(event) {
  var item = Office.context.mailbox.item;

  item.subject.getAsync(function (subjectResult) {
    if (subjectResult.status !== Office.AsyncResultStatus.Succeeded) {
      // Can't read subject — allow send
      event.completed({ allowEvent: true });
      return;
    }

    var subject = subjectResult.value || "";
    var hasCode = /^\[.+?\]/.test(subject.trim());

    if (hasCode) {
      // Subject already has a [CODE] prefix — send normally
      event.completed({ allowEvent: true });
      return;
    }

    // No code found — check if user already dismissed the reminder once
    item.loadCustomPropertiesAsync(function (propsResult) {
      if (propsResult.status !== Office.AsyncResultStatus.Succeeded) {
        event.completed({ allowEvent: true });
        return;
      }

      var props = propsResult.value;
      var bypass = props.get("bypassCodeReminder");

      if (bypass) {
        // User clicked Send a second time — allow through and clear the flag
        props.set("bypassCodeReminder", false);
        props.saveAsync(function () {
          event.completed({ allowEvent: true });
        });
        return;
      }

      // First attempt with no code — cancel send and show reminder
      props.set("bypassCodeReminder", true);
      props.saveAsync(function () {
        item.notificationMessages.addAsync("clientCodeReminder", {
          type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
          message: "Reminder: no client code prefix detected. Add [CODE] to the subject, or click Send again to send without one.",
          icon: "Icon.16x16",
          persistent: false,
        }, function () {
          event.completed({ allowEvent: false });
        });
      });
    });
  });
}
