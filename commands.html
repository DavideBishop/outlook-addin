Office.initialize = function () {};

function onItemSend(event) {
  var item = Office.context.mailbox.item;

  item.subject.getAsync(function (subjectResult) {
    if (subjectResult.status !== Office.AsyncResultStatus.Succeeded) {
      event.completed({ allowEvent: true });
      return;
    }

    var subject = subjectResult.value || "";
    var hasCode = /^\[.+?\]/.test(subject.trim());

    if (hasCode) {
      event.completed({ allowEvent: true });
      return;
    }

    item.loadCustomPropertiesAsync(function (propsResult) {
      if (propsResult.status !== Office.AsyncResultStatus.Succeeded) {
        event.completed({ allowEvent: true });
        return;
      }

      var props = propsResult.value;
      var bypass = props.get("bypassCodeReminder");

      if (bypass) {
        props.set("bypassCodeReminder", false);
        props.saveAsync(function () {
          event.completed({ allowEvent: true });
        });
        return;
      }

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
