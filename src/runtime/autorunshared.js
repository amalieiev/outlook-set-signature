// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// Contains code for event-based activation on both Outlook on web and Outlook on Windows.

/**
 * Gets template name (A,B,C) mapped based on the compose type
 * @param {*} compose_type The compose type (reply, forward, newMail)
 * @returns Name of the template to use for the compose type
 */
function get_template_name(compose_type) {
    if (compose_type === "reply")
        return Office.context.roamingSettings.get("reply");
    if (compose_type === "forward")
        return Office.context.roamingSettings.get("forward");
    return Office.context.roamingSettings.get("newMail");
}

function insertSignature(eventObj) {
    const signature = `
    <table>
      <tr>
        <td>Name</td>
        <td>Position</td>
      </tr>
    </table>
    `;

    Office.context.mailbox.item.body.setSignatureAsync(
        signature,
        {
            coercionType: "html",
            asyncContext: eventObj,
        },
        function (asyncResult) {
            asyncResult.asyncContext.completed();
        }
    );
}

Office.actions.associate("checkSignature", insertSignature);
