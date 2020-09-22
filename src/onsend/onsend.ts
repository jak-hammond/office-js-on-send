import getGlobal from '../getGlobal';
/* global , Office, console, setTimeout, Promise */

Office.onReady(() => {
    // If needed, Office.js is ready to be called
  });

let sendEvent: Office.AddinCommands.Event;

function processOnSendEvent(event: Office.AddinCommands.Event) {    
    console.info('Pausing send event...');
    sendEvent = event;

    setTimeout(async () => {
        try {
            // getAndremoveRecipients();
            await setInternetHeaders(); 

            // This will not add the header
            //await saveMessage();

            // This will add the header
            await patchBody();      

            sendEvent.completed({ allowEvent: true });
        } catch (err) {
            console.error(err);
            sendEvent.completed({ allowEvent: false });
        }
    }, 5000);
}

function saveMessage(): Promise<string> {
    return new Promise((resolve, reject) => {
        Office.context.mailbox.item.saveAsync((asyncResult: Office.AsyncResult<string>) => {
            if(asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error("Saving message failed.", asyncResult.error);
                reject(asyncResult.error);
            } else {
                resolve(asyncResult.value);
            }
        });
    });
}

function patchBody(): Promise<void> {
    return new Promise((resolve, reject) => {
        Office.context.mailbox.item.body.prependAsync("", {
            coercionType: Office.CoercionType.Text
        }, (asyncResult: Office.AsyncResult<void>) => {
            if(asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error("Failed prepending message body.", asyncResult.error);
                reject(asyncResult.error);
            } else {
                resolve();
            }
        });
    });
}

function setInternetHeaders(): Promise<void> {
    return new Promise((resolve, reject) => {
        Office.context.mailbox.item.internetHeaders.setAsync({
            "x-onsend-header": "bar"
          }, (asyncResult: Office.AsyncResult<void>) => {
              if(asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error("Failed setting internet message header through onsend function file.", asyncResult.error);
                reject(asyncResult.error);
              } else {
                console.info("Successfully set internet message header through onsend function file.");
                resolve();
              }
          });
    });
}

// function getAndremoveRecipients() {
//     Office.context.mailbox.item.to.getAsync((asyncResult: Office.AsyncResult<Office.EmailAddressDetails[]>) => {
//         if (asyncResult.error) {
//             console.error(asyncResult.error);
//         } else {
//             console.info('Removing recipients!');
//             const recipients = asyncResult.value;
//             // removeRecipients(recipients, () => {
//             //     console.info('Recipients removed!');
//             //     sendEvent.completed({ allowEvent: false });
//             // });

//             removeOneRecipient(recipients);
//         }
//     });
// }

// function removeOneRecipient(recipients:Office.EmailAddressDetails[]) {
//     const recipient = recipients[0];
//     const filtered = recipients.filter(x => x.emailAddress !== recipient.emailAddress);
//     Office.context.mailbox.item.to.setAsync(filtered, (asyncResult:Office.AsyncResult<void>) => {
//         if (asyncResult.error) {
//             console.error(asyncResult.error);
//         } else {
//             console.info('Recipient removed!');
//             sendEvent.completed({ allowEvent: false });
//         }
//     });
// }

// function removeRecipients(recipients:Office.EmailAddressDetails[], callback:Function) {
//     let recipient = recipients[0];
//     const filtered = recipients.filter(x => x.emailAddress !== recipient.emailAddress);
//     Office.context.mailbox.item.to.setAsync(filtered, (asyncResult:Office.AsyncResult<void>) => {
//         if (asyncResult.error) {
//             console.error(asyncResult.error);
//         } else {
//             console.info('Removed recipient!', recipient.emailAddress);
//             if (filtered.length) {
//                 removeRecipients(filtered, callback);
//             } else {
//                 callback();
//             }
//         }
//     });
// }

let g = getGlobal() as any;

// the add-in command functions need to be available in global scope
g.processOnSendEvent = processOnSendEvent;