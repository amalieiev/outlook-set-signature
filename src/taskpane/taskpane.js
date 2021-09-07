Office.initialize = function () {
    console.log("initialize");
};

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
        },
        function () {
            console.log("inserted");
        }
    );
}

export const getGlobal = () => {
    if (typeof self !== "undefined") {
        return self;
    }
    if (typeof window !== "undefined") {
        return window;
    }
    return typeof global !== "undefined" ? global : undefined;
};

const g = getGlobal();

g.insertSignature = insertSignature;
