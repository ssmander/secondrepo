/*
* Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
* See LICENSE in the project root for license information.
*/

function onMessageSendHandler(event) {
    Office.context.mailbox.item.sensitivityLabel.getAsync(
        { asyncContext: event },
        getLabelCallback
    );
}

function getLabelCallback(asyncResult) {
    const event = asyncResult.asyncContext;
    let label = "";
    const sensitiveEmailContent = "<p>This is a confidential email. Its contents are encrypted xml file</p>"
    if (asyncResult.status !== Office.AsyncResultStatus.Failed && asyncResult.value !== undefined) {
        label = asyncResult.value;
        var foundLabel = false;
        getSensitivityLabels().then(response => {
            response.sensitivityLabels.forEach((sensitivityLabel) => {
                if(sensitivityLabel.ID == label) {
                    if(sensitivityLabel.Name == "test") {
                        foundLabel = true;
                        Office.context.mailbox.item.body.prependAsync(sensitiveEmailContent, { coercionType: Office.CoercionType.Html },
                            function (result) {
                              if (result.status === Office.AsyncResultStatus.Failed) {
                                const message = "Failed to insert security text";
                                console.error(message);
                                event.completed({ allowEvent: false, errorMessage: message });
                                return;
                              }
                              else {
                                event.completed({ allowEvent: true });
                                return;
                              }
                            }
                        );
                    }
                }
            });

            if(!foundLabel) {
                event.completed({ allowEvent: true });
                return;
            }
        });
    } else {
        const message = "Failed to get sensitivity label";
        console.error(message);
        event.completed({ allowEvent: false, errorMessage: message });
        return;
    }

}

async function getSensitivityLabels() {
    return new Promise((resolve, reject) => {
      Office.context.sensitivityLabelsCatalog.getIsEnabledAsync((asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded && asyncResult.value == true) {
          // Identify available sensitivity labels in the catalog.
          Office.context.sensitivityLabelsCatalog.getAsync((asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
              const catalog = asyncResult.value;
              var sensitivityLabelsArray = new Array();
              catalog.forEach((sensitivityLabel) => {
                if(sensitivityLabel.children != null) {
                  sensitivityLabel.children.forEach((childLabel) => {
                    sensitivityLabelsArray.push({
                      Name: sensitivityLabel.name.trim() + " - " + childLabel.name.trim(),
                      ID: childLabel.id
                    });
                  });
                }
                else {
                  sensitivityLabelsArray.push({
                    Name: sensitivityLabel.name.trim(),
                    ID: sensitivityLabel.id
                  });
                }
              });
              resolve({
                name: "OK",
                sensitivityLabels: sensitivityLabelsArray
              });
            } else {
              reject({
                name: "Failed",
                message: "Action failed with error: " + asyncResult.error.message
              });
            }
          });
        } else {
          reject({
            name: "Failed",
            message: "Action failed with error: " + asyncResult.error.message
          });
        }
      });
    });
  }

// IMPORTANT: To ensure your add-in is supported in the Outlook client on Windows, remember to map the event handler name specified in the manifest to its JavaScript counterpart.
if (Office.context.platform === Office.PlatformType.PC || Office.context.platform == null) {
    Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
}