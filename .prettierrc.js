/**
 * @type {import("prettier").Config}
 */
export default {
    tabWidth: 4,
    printWidth: 100,
    trailingComma: "es5",
    overrides: [
        {
            files: "*.json",
            options: {
                tabWidth: 2,
            },
        },
    ],
};
