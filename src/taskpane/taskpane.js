/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

// var messageBanner;

// Office.onReady((info) => {
//   if (info.host === Office.HostType.Outlook) {
//     document.getElementById("sideload-msg").style.display = "none";
//     // document.getElementById("app-body").style.display = "flex";
//     // document.getElementById("run").onclick = run;

//     var element = document.querySelector('.ms-MessageBanner');
//     messageBanner = new app.notification.MessageBanner(element);
//     messageBanner.hideBanner();
//   }
// });

// export async function run() {
//   /**
//    * Insert your Outlook code here
//    */
// }

// function errorHandler(error) {
//   showNotification(error);
// }

// // Display notifications in message banner at the top of the task pane.
// function showNotification(content) {
//   $("#notificationBody").text(content);
//   messageBanner.showBanner();
//   messageBanner.toggleExpansion();
// }