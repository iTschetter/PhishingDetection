/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  /**
   * Insert your Outlook code here
   */

  //const item = Office.context.mailbox.item;
  Office.context.mailbox.item.body.getAsync(
    "text",
    { asyncContext: "This is passed to the callback" },
    function callback(result) {
      // Do something with the result.

      let insertAt = document.getElementById("item-subject");
      let label = document.createElement("b").appendChild(document.createTextNode("Subject: "));
      insertAt.appendChild(label);
      insertAt.appendChild(document.createElement("br"));
      insertAt.appendChild(document.createTextNode(result.value));
      insertAt.appendChild(document.createElement("br"));
    }
  );
}
