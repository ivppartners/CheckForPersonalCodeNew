function onMessageSendHandler(event) {
    Office.context.mailbox.item.body.getAsync(
      "text",
      { asyncContext: event },
      getBodyCallback
    );
  }
  
  function getBodyCallback(asyncResult){
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
  
    const matches = containsLithuanianPersonalCode(body);
    if (matches) {
        event.completed({
            allowEvent: false,
            errorMessage: "Panašu, kad į laiško tekstą įtraukėte asmens kodą.",
            errorMessageMarkdown: "Panašu, kad į laiško tekstą įtraukėte asmens kodą.",
            sendModeOverride: Office.MailboxEnums.SendModeOverride.PromptUser
          });

    } else {
      event.completed({ allowEvent: true });
    }
  }

  function containsLithuanianPersonalCode(text) {
    if (!text || typeof text !== "string") {
        return false;
    }

    // Regular expression to match 11-digit numbers
    const personalCodeRegex = /\b\d{11}\b/g;
    const matches = text.match(personalCodeRegex);

    if (!matches) {
        return false;
    }

    // Validate each match using Lithuanian personal code rules
    for (const code of matches) {
        if (isValidLithuanianPersonalCode(code)) {
            return true;
        }
    }

    return false;
}

function isValidLithuanianPersonalCode(code) {
    if (code.length !== 11) {
        return false;
    }

    // First digit indicates gender and century
    const firstDigit = parseInt(code[0], 10);
    if (firstDigit < 1 || firstDigit > 6) {
        return false;
    }

    // Validate checksum
    const weights1 = [1, 2, 3, 4, 5, 6, 7, 8, 9, 1];
    const weights2 = [3, 4, 5, 6, 7, 8, 9, 1, 2, 3];

    const digits = code.split("").map(Number);
    const checksum = digits[10];

    const sum1 = digits.slice(0, 10).reduce((sum, digit, index) => sum + digit * weights1[index], 0);
    const remainder1 = sum1 % 11;

    if (remainder1 !== 10) {
        return remainder1 === checksum;
    }

    const sum2 = digits.slice(0, 10).reduce((sum, digit, index) => sum + digit * weights2[index], 0);
    const remainder2 = sum2 % 11;

    return remainder2 === checksum;
  }
  
  // IMPORTANT: To ensure your add-in is supported in Outlook, remember to map the event handler name specified in the manifest to its JavaScript counterpart.
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);