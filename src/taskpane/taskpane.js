import { insertSignature } from "../helpers/office";

Office.initialize = function () {
    console.log("initialize");
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

g.insertSignature = insertSignature;
