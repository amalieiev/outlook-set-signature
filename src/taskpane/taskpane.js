// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

let _user_info;

Office.initialize = function (reason) {
    console.log("initialize!!!");
    // on_initialization_complete();
};

function on_initialization_complete() {
    $(document).ready(function () {
        lazy_init_user_info();
        populate_templates();
        show_signature_settings();
    });
}

function lazy_init_user_info() {
    if (!_user_info) {
        let user_info_str = localStorage.getItem("user_info");

        if (user_info_str) {
            _user_info = JSON.parse(user_info_str);
        } else {
            console.log("Unable to retrieve 'user_info' from localStorage.");
        }
    }
}

function populate_templates() {
    populate_template_A();
    populate_template_B();
    populate_template_C();
}

function populate_template_A() {
    let str = get_template_A_str(_user_info);
    $("#box_1").html(str);
}

function populate_template_B() {
    let str = get_template_B_str(_user_info);
    $("#box_2").html(str);
}

function populate_template_C() {
    let str = get_template_C_str(_user_info);
    $("#box_3").html(str);
}

function show_signature_settings() {
    let val = Office.context.roamingSettings.get("newMail");
    if (val) {
        $("#new_mail").val(val);
    }

    val = Office.context.roamingSettings.get("reply");
    if (val) {
        $("#reply").val(val);
    }

    val = Office.context.roamingSettings.get("forward");
    if (val) {
        $("#forward").val(val);
    }

    val = Office.context.roamingSettings.get("override_olk_signature");
    if (val != null) {
        $("#checkbox_sig").prop("checked", val);
    }
}

// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

function save_user_settings_to_roaming_settings() {
    Office.context.roamingSettings.saveAsync(function (asyncResult) {
        console.log(
            "save_user_info_str_to_roaming_settings - " +
                JSON.stringify(asyncResult)
        );
    });
}

function disable_client_signatures_if_necessary() {
    if ($("#checkbox_sig").prop("checked") === true) {
        Office.context.mailbox.item.disableClientSignatureAsync(function (
            asyncResult
        ) {
            console.log(
                "disable_client_signature_if_necessary - " +
                    JSON.stringify(asyncResult)
            );
        });
    }
}

function save_signature_settings() {
    let user_info_str = localStorage.getItem("user_info");

    if (user_info_str) {
        if (!_user_info) {
            _user_info = JSON.parse(user_info_str);
        }

        Office.context.roamingSettings.set("user_info", user_info_str);
        Office.context.roamingSettings.set(
            "newMail",
            $("#new_mail option:selected").val()
        );
        Office.context.roamingSettings.set(
            "reply",
            $("#reply option:selected").val()
        );
        Office.context.roamingSettings.set(
            "forward",
            $("#forward option:selected").val()
        );

        Office.context.roamingSettings.set(
            "override_olk_signature",
            $("#checkbox_sig").prop("checked")
        );

        save_user_settings_to_roaming_settings();

        disable_client_signatures_if_necessary();

        $("#message").show("slow");
    } else {
        // TBD display an error somewhere?
    }
}

function set_body(str) {
    Office.context.mailbox.item.body.setAsync(
        get_cal_offset() + str,

        {
            coercionType: Office.CoercionType.Html,
        },

        function (asyncResult) {
            console.log("set_body - " + JSON.stringify(asyncResult));
        }
    );
}

function set_signature(str) {
    Office.context.mailbox.item.body.setSignatureAsync(
        str,

        {
            coercionType: Office.CoercionType.Html,
        },

        function (asyncResult) {
            console.log("set_signature - " + JSON.stringify(asyncResult));
        }
    );
}

function insert_signature(str) {
    if (
        Office.context.mailbox.item.itemType ==
        Office.MailboxEnums.ItemType.Appointment
    ) {
        set_body(str);
    } else {
        set_signature(str);
    }
}

function test_template_A() {
    let str = get_template_A_str(_user_info);
    console.log("test_template_A - " + str);

    insert_signature(str);
}

function test_template_B() {
    let str = get_template_B_str(_user_info);
    console.log("test_template_B - " + str);

    insert_signature(str);
}

function test_template_C() {
    let str = get_template_C_str(_user_info);
    console.log("test_template_C - " + str);

    insert_signature(str);
}

function insertSignature() {
    console.log("insert signature");

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
            // asyncContext: eventObj,
        },
        function (asyncResult) {
            console.log("inserted");
            // asyncResult.asyncContext.completed();
        }
    );
}
