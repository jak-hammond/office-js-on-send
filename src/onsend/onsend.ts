import getGlobal from '../getGlobal';
/* eslint-disable */
/* global , Office */

Office.onReady(() => {
    // If needed, Office.js is ready to be called
  });

let sendEvent: Office.AddinCommands.Event;

function processOnSendEvent(event: Office.AddinCommands.Event) {    
    console.info('Pausing send event');
    sendEvent = event;

    setTimeout(() => {
        getAndremoveRecipients();
    }, 5000);
}

function getAndremoveRecipients() {
    Office.context.mailbox.item.to.getAsync((asyncResult: Office.AsyncResult<Office.EmailAddressDetails[]>) => {
        if (asyncResult.error) {
            console.error(asyncResult.error);
        } else {
            console.info('Removing recipients!');
            const recipients = asyncResult.value;
            // removeRecipients(recipients, () => {
            //     console.info('Recipients removed!');
            //     sendEvent.completed({ allowEvent: false });
            // });

            removeOneRecipient(recipients);
        }
    });
}

function removeOneRecipient(recipients:Office.EmailAddressDetails[]) {
    const recipient = recipients[0];
    const filtered = recipients.filter(x => x.emailAddress !== recipient.emailAddress);
    Office.context.mailbox.item.to.setAsync(filtered, (asyncResult:Office.AsyncResult<void>) => {
        if (asyncResult.error) {
            console.error(asyncResult.error);
        } else {
            console.info('Recipient removed!');
            sendEvent.completed({ allowEvent: false });
        }
    });
}

function removeRecipients(recipients:Office.EmailAddressDetails[], callback:Function) {
    let recipient = recipients[0];
    const filtered = recipients.filter(x => x.emailAddress !== recipient.emailAddress);
    Office.context.mailbox.item.to.setAsync(filtered, (asyncResult:Office.AsyncResult<void>) => {
        if (asyncResult.error) {
            console.error(asyncResult.error);
        } else {
            console.info('Removed recipient!', recipient.emailAddress);
            if (filtered.length) {
                removeRecipients(filtered, callback);
            } else {
                callback();
            }
        }
    });
}

let g = getGlobal() as any;

// the add-in command functions need to be available in global scope
g.processOnSendEvent = processOnSendEvent;