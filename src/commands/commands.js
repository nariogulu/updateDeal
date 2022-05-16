/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window */

// The initialize function must be run each time a new page is loaded
(function () {
  Office.initialize = function (reason) {
    //If you need to initialize something you can do so here.
  };
})();

Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */
function action(event) {
  const message = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Performed action.",
    icon: "Icon.80x80",
    persistent: true,
  };

  // Show a notification message
  Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);

  // Be sure to indicate when the add-in command function is complete
  event.completed();
}

function update(event) {
  //Consult Office.js API reference to see all you can do. This just shows the simplest action.

  Word.run(async (context) => {
    /**
     * Insert your Word code here
     *
     */
    const url = "https://jsonplaceholder.typicode.com/todos/1";
    const response = await fetch(url);

    //Expect that status code is in 200-299 range
    if (!response.ok) {
      throw new Error(response.statusText);
    }
    const data = await response.json();
    console.log("Response: " + data.id);

    let doc = context.document;
    let paragraphs = context.document.body.paragraphs;
    paragraphs.load("$none");

    await context.sync();

    let contentControls = context.document.contentControls.getByTag("amount");
    contentControls.load("text");

    await context.sync();

    //copied from docs, will refactor
    let a = Math.random();
    for (let i = 0; i < contentControls.items.length; i++) {
      contentControls.items[i].insertText((data.id * a * 1000).toString(), "Replace");
    }

    await context.sync();
  });
  //Required, call event.completed to let the platform know you are done processing.
  event.completed();
}

function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
    ? window
    : typeof global !== "undefined"
    ? global
    : undefined;
}

const g = getGlobal();

// The add-in command functions need to be available in global scope
g.action = update;
g.update = update;
