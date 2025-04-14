/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    // document.getElementById("add-attachment").onclick = addAttachment;
    // document.getElementById("add-inline-image").onclick = addInlineImage;
  }
});

// Adds the specified URL as an attachment to the message.
// Adds the specified URL as an attachment to the message.
// Then, programmatically sends the message.
// export async function addAttachment() {
//   const attachmentUrl = document.querySelector("#attachment-url").value;
//   Office.context.mailbox.item.addFileAttachmentAsync(attachmentUrl, getFileName(attachmentUrl), (asyncResult) => {
//     console.log("Added attachment");

//     // Programmatically send the item.
//     Office.context.mailbox.item.sendAsync((sendResult) => {
//         if (sendResult.status === Office.AsyncResultStatus.Failed) {
//             console.log(`Unable to send the message: ${sendResult.error.message}`);
//             return;
//         }

//         // Don't run additional operations here. Process them before the sendAsync call instead.
//     });

//     // Don't run additional operations after the sendAsync call. Process them before the sendAsync call instead.
//   });
// }

// Gets the file name from a URL.
// function getFileName(url) {
//   const lastIndex = url.lastIndexOf("/");
//   if (lastIndex >= 0) {
//     return url.substring(lastIndex + 1);
//   }

//   return url;
// }

// Adds an inline image to the body of the message.
// export async function addInlineImage() {
//   const mailItem = Office.context.mailbox.item;
//   const base64String =
//     "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAnUExURQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAN0S+bUAAAAMdFJOUwAQIDBAUI+fr7/P7yEupu8AAAAJcEhZcwAADsMAAA7DAcdvqGQAAAF8SURBVGhD7dfLdoMwDEVR6Cspzf9/b20QYOthS5Zn0Z2kVdY6O2WULrFYLBaLxd5ur4mDZD14b8ogWS/dtxV+dmx9ysA2QUj9TQRWv5D7HyKwuIW9n0vc8tkpHP0W4BOg3wQ8wtlvA+PC1e8Ao8Ld7wFjQtHvAiNC2e8DdqHqKwCrUPc1gE1AfRVgEXBfB+gF0lcCWoH2tYBOYPpqQCNwfT3QF9i+AegJfN8CtAWhbwJagtS3AbIg9o2AJMh9M5C+SVGBvx6zAfmT0r+Bv8JMwP4kyFPir+cswF5KL3WLv14zAFBCLf56Tw9cparFX4upgaJUtPhrOS1QlY5W+vWTXrGgBFB/b72ev3/0igUdQPppP/nfowfKUUEFcP207y/yxKmgAYQ+PywoAFOfCH3A2MdCFzD3kdADBvq10AGG+pXQBgb7pdAEhvuF0AIc/VtoAK7+JciAs38KIuDugyAC/v4hiMCE/i7IwLRBsh68N2WQjMVisVgs9i5bln8LGScNcCrONQAAAABJRU5ErkJggg==";

//   // Get the current body of the message.
//   mailItem.body.getAsync(Office.CoercionType.Html, (bodyResult) => {
//     if (bodyResult.status === Office.AsyncResultStatus.Failed) {
//       console.log(bodyResult.error.message);
//       return;
//     }

//     // Insert the Base64-encoded image to the beginning of the body.
//     const options = { isInline: true, asyncContext: bodyResult.value };
//     mailItem.addFileAttachmentFromBase64Async(base64String, "sample.png", options, (attachResult) => {
//       if (attachResult.status === Office.AsyncResultStatus.Failed) {
//         console.log(attachResult.error.message);
//         return;
//       }

//       let body = attachResult.asyncContext;
//       body += '<img src="cid:sample.png" />';
//       mailItem.body.setAsync(body, { coercionType: Office.CoercionType.Html }, (setResult) => {
//         if (setResult.status === Office.AsyncResultStatus.Failed) {
//           console.log(setResult.error.message);
//           return;
//         }

//         console.log("Inline image added to the body.");
//       });
//     });
//   });
// }