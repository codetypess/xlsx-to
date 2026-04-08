const traces: string[] = [];

export const trace = (msg: string) => {
    traces.push(msg);
    return new (class {
        [Symbol.dispose]() {
            traces.pop();
        }
    })();
};

export function error(msg: string): never {
    let str = "";
    if (traces.length > 0) {
        str = "\n" + traces.map((v) => `    -> ${v}`).join("\n");
    }
    throw new Error(msg + str);
}

export function assert(condition: unknown, msg: string): asserts condition {
    if (!condition) {
        error(msg);
    }
}
