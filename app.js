const ALLOWED_DOMAIN = "test.dd";
const NOTIFICATION_ID = "externalRecipientsWarning";

Office.onReady(() => {
    window.onRecipientsChanged = onRecipientsChanged;
    window.onMessageSend = onMessageSend;
});

function onRecipientsChanged(event) {
    checkRecipients()
        .then(external => {
            if (external.length > 0) {
                showWarning(external);
            } else {
                clearWarning();
            }
        })
        .finally(() => event.completed());
}

function onMessageSend(event) {
    checkRecipients()
        .then(external => {
            if (external.length > 0) {
                const emails = external.map(e => e.emailAddress).join(", ");
                event.completed({
                    allowEvent: false,
                    errorMessage: `Nie można wysłać maila. Adresaci spoza domeny: ${emails}`
                });
            } else {
                event.completed({ allowEvent: true });
            }
        })
        .catch(() => {
            event.completed({ allowEvent: true });
        });
}

function checkRecipients() {
    const item = Office.context.mailbox.item;

    return Promise.all([
        getRecipients(item.to),
        getRecipients(item.cc),
        getRecipients(item.bcc)
    ])
    .then(([to, cc, bcc]) => {
        const all = [...to, ...cc, ...bcc];
        return all.filter(r => isExternal(r.emailAddress));
    });
}

function getRecipients(field) {
    return new Promise((resolve) => {
        field.getAsync(result => {
            resolve(result.value || []);
        });
    });
}

function isExternal(email) {
    if (!email) return false;
    return !email.toLowerCase().endsWith("@" + ALLOWED_DOMAIN);
}

function showWarning(externals) {
    const emails = externals.map(e => e.emailAddress).join(", ");

    Office.context.mailbox.item.notificationMessages.replaceAsync(NOTIFICATION_ID, {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: `Uwaga! Adresaci spoza domeny: ${emails}`,
        persistent: true
    });
}

function clearWarning() {
    Office.context.mailbox.item.notificationMessages.removeAsync(NOTIFICATION_ID);
}
