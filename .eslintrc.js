/**
 * @see https://github.com/typescript-eslint/typescript-eslint/blob/main/packages/eslint-plugin/docs/rules/naming-convention.md
 */

module.exports = {
	env: {
		browser: true,
		es2021: true,
		node: true,
	},
	extends: [
		"plugin:react/recommended",
		"standard-with-typescript",
		"plugin:@typescript-eslint/recommended",
	],
	overrides: [],
	parserOptions: {
		ecmaVersion: "latest",
		sourceType: "module",
		project: ["./tsconfig.json"],
	},
	plugins: ["react", "@typescript-eslint"],
	ignorePatterns: [".eslintrc.js", "*.mjs", "demos/*", "index.d.ts", "gulpfile.js"],
	rules: {
		"@typescript-eslint/indent": ["error", "tab"],
		"@typescript-eslint/prefer-nullish-coalescing": 0, // "warn", too many items!
		"@typescript-eslint/restrict-plus-operands": "warn", // TODO: "error"
		"@typescript-eslint/restrict-template-expressions": "warn", // TODO: "error"
		"@typescript-eslint/strict-boolean-expressions": 0,
		"comma-dangle": ["error", "only-multiline"],
		"no-lone-blocks": 0,
		"no-tabs": ["error", { allowIndentationTabs: true }],
		indent: ["error", "tab", { "SwitchCase": 1, "ImportDeclaration": 1 }],
		quotes: ["error", "single"],
		semi: ["error", "never"],
	},
};
