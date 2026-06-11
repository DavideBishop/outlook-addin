Office.onReady();

function onItemSend(event) {
  Office.context.mailbox.item.subject.getAsync(function(result) {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      event.completed({ allowEvent: true });
      return;
    }

    var subject = (result.value || "").trim();
    var hasCode = /^\[.+?\]/.test(subject);

    if (hasCode) {
      event.completed({ allowEvent: true });
    } else {
      event.completed({
        allowEvent: false,
        errorMessage: "Reminder: no client code detected. Add [CODE] to the start of the subject, or click Send Anyway to send without one."
      });
    }
  });
}
