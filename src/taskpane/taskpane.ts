/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady(info => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  Office.context.mailbox.item.addHandlerAsync(Office.EventType.RecipientsChanged,
    () => {
      /* eslint-disable-next-line */
      console.info('Recipients have changed!');
    })
}
