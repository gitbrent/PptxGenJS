const CleanWebpackPlugin = require("clean-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");

module.exports = {
	entry: "./src/pptxgen.ts",
	mode: "development",
	//mode: "production", // these 2 made no diff in file size
	//devtool: false,
	output: {
		path: __dirname + "/dist",
		filename: "build/[name].[contenthash].js",
		publicPath: "/"
	},
	resolve: {
		extensions: [".ts", ".tsx", ".js", ".json"]
	},
	module: {
		rules: [
			{
				test: /\.js?$/,
				exclude: /(node_modules|bower_components)/,
				use: {
					loader: "babel-loader",
					options: {
						presets: ["@babel/preset-env"]
					}
				}
			},
			{
				test: /\.ts?$/,
				loader: "ts-loader"
			}
		]
	}
};
