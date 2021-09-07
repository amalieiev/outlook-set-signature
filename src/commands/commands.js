import { insertSignature } from "../helpers/office";

Office.onReady(() => {
    // Call to initialise the Office components and enable the event based function
});

const onNewMessageHandler = async (event) => {
    await insertSignature();
    event.completed();
};

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

g.onNewMessageHandler = onNewMessageHandler;

Office.actions.associate("onNewMessageHandler", onNewMessageHandler);
