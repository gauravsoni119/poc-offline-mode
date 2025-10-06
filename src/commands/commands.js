/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

Office.onReady(() => {
  // If needed, Office.js is ready to be called.
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

  // Show a notification message.
  Office.context.mailbox.item.notificationMessages.replaceAsync(
    "ActionPerformanceNotification",
    message
  );

  // Be sure to indicate when the add-in command function is complete.
  event.completed();
}

// Register the function with Office.
Office.actions.associate("action", action);

/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

function onMessageSendHandler(event) {
  Office.context.mailbox.item.body.getAsync("text", { asyncContext: event }, getBodyCallback);
}

function getBodyCallback(asyncResult) {
  console.info('Scanning the message body')
  const event = asyncResult.asyncContext;
  let body = "";
  if (asyncResult.status !== Office.AsyncResultStatus.Failed && asyncResult.value !== undefined) {
    body = asyncResult.value;
  } else {
    const message = "Failed to get body text";
    console.error(message);
    event.completed({ allowEvent: false, errorMessage: message });
    return;
  }

  const POC_login = ["login"];
  const matches_login = hasMatches(body, POC_login);

  const POC_warning = ["warning"];
  const matches_warning = hasMatches(body, POC_warning);

  const POC_away = ["away"];
  const matches_away = hasMatches(body, POC_away);

  const POC_long = ["long", "running"];
  const matches_long = hasMatches(body, POC_long);

  const POC_attachment = ["file"];
  const matches_attachment = hasMatches(body, POC_attachment);

  const POC_reuse = ["reuse"];
  const matches_reuse = hasMatches(body, POC_reuse);

  if (matches_login) {
    event.completed({
      allowEvent: false,
      errorMessage:
        "You need to be logged in with your account to protect your messages and files against possible data leaks",
      errorMessageMarkdown:
        "You need to be logged in with your account to protect your messages and files against possible data leaks.",
      cancelLabel: "Login",
      commandId: "msgComposeOpenPaneButton",
      contextData: {
        autoLogin: true,
      },
      sendModeOverride: Office.MailboxEnums.SendModeOverride.PromptUser,
    });
  } else if (matches_warning) {
    event.completed({
      allowEvent: false,
      errorMessage:
        "Warning — Email seems to contain a some sensitive data, for which your organization's policy recommends secure enable.",
      errorMessageMarkdown:
        "**Warning** — Email seems to contain a some sensitive data, for which your organization's policy recommends secure enable.",
      cancelLabel: "Resolve",
      commandId: "msgComposeOpenPaneButton",
      contextData: {
        autoViolation: true,
      },
    });
  } else if (matches_away) {
    setTimeout(() => {
      event.completed({
        allowEvent: true,
      });
    }, 4 * 1000);
  } else if (matches_long) {
    setTimeout(() => {
      event.completed({
        allowEvent: true,
      });
    }, 10 * 1000);
  } else if (matches_attachment) {
    event.completed({
      allowEvent: false,
      errorMessage:
        "Your Attachments — Wait for your attachments to upload before sending this email. This might take a few minutes.",
      errorMessageMarkdown:
        "**Your Attachments** — Wait for your attachments to upload before sending this email. This might take a few minutes.",
      cancelLabel: "Check attachments",
      commandId: "msgComposeOpenPaneButton",
      contextData: {
        autoAttachments: true,
      },
    });
  } else if (matches_reuse) {
    event.completed({
      allowEvent: false,
      errorMessage: "Validation issue — Please resolve all issue to send email",
      errorMessageMarkdown: "**Validation issue** — Please resolve all issue to send email",
      cancelLabel: "Resolve issues",
      commandId: "msgComposeOpenPaneButton",
      contextData: {
        autoOnSendDialog: true,
      },
    });
  } else {
    event.completed({ allowEvent: true });
  }
}

function hasMatches(body, terms) {
  if (body == null || body === "") {
    return false;
  }

  for (let index = 0; index < terms.length; index++) {
    const term = terms[index].trim();
    const regex = RegExp(term, "i");
    if (regex.test(body)) {
      return true;
    }
  }

  return false;
}

function getAttachmentsCallback(asyncResult) {
  const event = asyncResult.asyncContext;
  if (asyncResult.value.length > 0) {
    for (let i = 0; i < asyncResult.value.length; i++) {
      if (asyncResult.value[i].isInline == false) {
        event.completed({ allowEvent: true });
        return;
      }
    }

    event.completed({
      allowEvent: false,
      errorMessage:
        "Looks like the body of your message includes an image or an inline file. Attach a copy to the message before sending.",
      // TIP: In addition to the formatted message, it's recommended to also set a
      // plain text message in the errorMessage property for compatibility on
      // older versions of Outlook clients.
      errorMessageMarkdown:
        "Looks like the body of your message includes an image or an inline file. Attach a copy to the message before sending.\n\n**Tip**: For guidance on how to attach a file, see [Attach files in Outlook](https://www.contoso.com/help/attach-files-in-outlook).",
    });
  } else {
    event.completed({
      allowEvent: false,
      errorMessage: "Looks like you're forgetting to include an attachment.",
      // TIP: In addition to the formatted message, it's recommended to also set a
      // plain text message in the errorMessage property for compatibility on
      // older versions of Outlook clients.
      errorMessageMarkdown:
        "Looks like you're forgetting to include an attachment.\n\n**Tip**: For guidance on how to attach a file, see [Attach files in Outlook](https://www.contoso.com/help/attach-files-in-outlook).",
    });
  }
}

// IMPORTANT: To ensure your add-in is supported in Outlook, remember to map the event handler name specified in the manifest to its JavaScript counterpart.
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
