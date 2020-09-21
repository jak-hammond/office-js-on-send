/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, console */

Office.onReady(info => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  // Office.context.mailbox.item.addHandlerAsync(Office.EventType.RecipientsChanged,
  //   () => {      
  //     console.info('Recipients have changed!');
  //   });

  Office.context.mailbox.item.internetHeaders.setAsync({
    "x-taskpane-header": "foo"
  }, (asyncResult: Office.AsyncResult<void>) => {
      if(asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error("Failed setting internet message header through taskpane.", asyncResult.error);
      } else {
        console.info("Successfully set internet message header through taskpane.");
      }
  });
}
