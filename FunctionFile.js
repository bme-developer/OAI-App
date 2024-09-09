Office.onReady(function(info) {
    if (info.host === Office.HostType.Outlook) {
        // Office is ready
    }
});

// This function is called when a new message is composed
function onNewMessageComposeHandler(event) {
    // Insert "Hello" in the body of the email
    Office.context.mailbox.item.body.setAsync("Hello", { coercionType: Office.CoercionType.Html }, function(asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.error(asyncResult.error.message);
        }
    });

    // Call event.completed to indicate event processing is complete
    event.completed();
}
