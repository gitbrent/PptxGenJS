// [REF](https://github.com/typescript-eslint/typescript-eslint/blob/main/packages/eslint-plugin/docs/rules/naming-convention.md)
// https://dev.to/knowankit/setup-eslint-and-prettier-in-react-app-357b


module.exports = {
	env: {
		browser: true,
		es2021: true,
	},
	extends: ["eslint:recommended", "plugin:react/recommended", "plugin:@typescript-eslint/recommended"],
	parser: "@typescript-eslint/parser",
	parserOptions: {
		ecmaFeatures: {
			jsx: true,
		},
		ecmaVersion: "latest",
		sourceType: "module",
		project: ["./tsconfig.json"],
	},
	plugins: ["react", "@typescript-eslint"],
	ignorePatterns: [".eslintrc.js"],
	rules: {
		indent: ["error", "tab"],
		"linebreak-style": ["error", "unix"],
		quotes: ["error", "single"],
		semi: ["error", "never"],
		/** `indent` rules to match prettier */
		indent: ["error", "tab", { "SwitchCase": 2, "ImportDeclaration": 1 }],
		"@typescript-eslint/naming-convention": [
			"warn",
			/*{
				selector: "interface",
				format: ["PascalCase"],
				prefix: ["I"],
			},*/
			/* TODO:
			{
				selector: "variable",
				modifier: "const",
				format: ["PascalCase"],
			},
			*/
			{
				selector: "variable",
				format: ["camelCase", "UPPER_CASE"],
			},
			{
				selector: "variable",
				types: ["boolean"],
				format: ["camelCase", "UPPER_CASE"],
				prefix: ["is", "should", "has", "can", "did", "will"],
			},
		],
	},
};
