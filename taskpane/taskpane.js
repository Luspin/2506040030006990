/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

let listItems;

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    console.log("Office is ready for Outlook");

    let officeLoadedMessage = document.createElement("div");
    officeLoadedMessage.innerHTML = `<p>Office.js loaded.</p>`;
    document.body.appendChild(officeLoadedMessage);

    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    listItems = document.getElementById("selected-items");

    // Register an event handler to identify when messages are selected.
    Office.context.mailbox.addHandlerAsync(
      Office.EventType.SelectedItemsChanged,
      selectedItemsChangedHandler,
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.log(asyncResult.error.message);
          return;
        }

        console.log("Registered: 'selectedItemsChangedHandler'");
      });

    // Register an event handler to identify when messages are selected.
    Office.context.mailbox.addHandlerAsync(
      Office.EventType.ItemChanged,
      itemChangedHandler,
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.log(asyncResult.error.message);
          return;
        }

        console.log("Registered: 'itemChangedHandler'");
      });

  }
});


async function selectedItemsChangedHandler() {
  console.log("Invoked: 'selectedItemsChangedHandler'");
  // Clear the listItems variable.
  clearList(listItems);

  // Get the subject line and sender's email address of each selected message and log them to a list in the task pane.
  Office.context.mailbox.getSelectedItemsAsync((asyncResult) => {
    // Error Handler
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      console.log("ERROR: " + asyncResult.error.message);
      return;
    }

    console.log(`'selectedItemsChangedHandler' | Number of Selected Items: ${asyncResult.value.length}`);
    // Log each item to to the Console.
    asyncResult.value.forEach((message) => {
      console.log(`ItemID: ${message.itemId}`);
      console.log(`ConversationId: ${message.conversationId}`);
      console.log(`InternetMessageId: ${message.internetMessageId}`);
      console.log(`Subject: ${message.subject}`);
      console.log(`ItemType: ${message.itemType}`);
      console.log(`ItemMode: ${message.itemMode}`);
      console.log(`HasAttachment: ${message.hasAttachment}`);
      console.log("\n");
      addItemToList(message.subject, listItems);
    });
    getItemInfo('selectedItemsChangedHandler', asyncResult.value);
  });
}

async function itemChangedHandler() {
  console.log("Invoked: 'itemChangedHandler'");
  // Clear the listItems variable.
  clearList(listItems);

  Office.context.mailbox.getSelectedItemsAsync((asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      console.log(asyncResult.error.message);
      return;
    }

    console.log(`'itemChangedHandler' | Number of Selected Items: ${asyncResult.value.length}`);
    // Log each item to to the Console.
    asyncResult.value.forEach((message) => {
      console.log(`ItemID: ${message.itemId}`);
      console.log(`ConversationId: ${message.conversationId}`);
      console.log(`InternetMessageId: ${message.internetMessageId}`);
      console.log(`Subject: ${message.subject}`);
      console.log(`ItemType: ${message.itemType}`);
      console.log(`ItemMode: ${message.itemMode}`);
      console.log(`HasAttachment: ${message.hasAttachment}`);
      console.log("\n");
      addItemToList(message.subject, listItems);
    });
    getItemInfo('itemChangedHandler', asyncResult.value);
  });

}

// Gets the subject line and sender's email address of each selected message.
async function getItemInfo(invoker, selectedItems) {
  for (const item of selectedItems) {
    // Log the item to the Console.
    console.log("-- ITEM INFO");
    console.log(item);

    // retrieve item details
    let itemDetails = await Promise.resolve(Office.context.mailbox.item);
    console.log(itemDetails);

    addItemToList((invoker + " | " + itemDetails.subject), listItems);
  }
}

// Clears a List variable
async function clearList(list) {
  while (list.firstChild) {
    list.removeChild(list.firstChild);
  }
}

// Adds an Item to a List variable
async function addItemToList(item, list) {
  // console.trace("TRACE: Adding Item to List:", item);
  const listItem = document.createElement("li");
  listItem.textContent = item;
  list.appendChild(listItem);
}
