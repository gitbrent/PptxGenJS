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
	ignorePatterns: [".eslintrc.js"],
	rules: {
		"strict-boolean-expressions": 2,
		"comma-dangle": [
			"error",
			{
				arrays: "never",
				exports: "never",
				functions: "never",
				imports: "always",
				objects: "always",
			},
		],
		"no-lone-blocks": 0,
		"no-tabs": ["error", { allowIndentationTabs: true }],
		"@typescript-eslint/indent": ["error", "tab"],
		indent: ["error", "tab", { "SwitchCase": 2, "ImportDeclaration": 1 }],
		//indent: ["error", "tab"],
		//indent: ["error", "tab", { "SwitchCase": 2, "ImportDeclaration": 1 }],
		quotes: ["error", "single"],
		semi: ["error", "never"],
	},
};
