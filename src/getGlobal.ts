/* global , self, window, global */

function getGlobal() {
    return typeof self !== "undefined"
        ? self
        : typeof window !== "undefined"
        ? window
        : typeof global !== "undefined"
        ? global
        : undefined;
}

export default getGlobal;
