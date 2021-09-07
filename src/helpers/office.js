export function insertSignature() {
    const signature = `
    <table>
      <tr>
        <td>Signature will be displayed here</td>
      </tr>
    </table>
    `;

    return new Promise((resolve) => {
        Office.context.mailbox.item.body.setSignatureAsync(
            signature,
            {
                coercionType: "html",
            },
            function () {
                resolve();
            }
        );
    });
}
