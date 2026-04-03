const doings: string[] = [];

export const doing = (msg: string) => {
    doings.push(msg);
    return new (class {
        [Symbol.dispose]() {
            doings.pop();
        }
    })();
};

export function error(msg: string): never {
    let str = "";
    if (doings.length > 0) {
        str = "\n" + doings.map((v) => `    -> ${v}`).join("\n");
    }
    throw new Error(msg + str);
}

export function assert(condition: unknown, msg: string): asserts condition {
    if (!condition) {
        error(msg);
    }
}
