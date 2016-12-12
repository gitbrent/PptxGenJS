var webpack = require('webpack'),
    UglifyJsPlugin = webpack.optimize.UglifyJsPlugin,
    path = require('path'),
    //env = require('yargs').argv.mode,
    env = '',//process.env.WEBPACK_ENV,
    plugins = [],
    outputFile;

var  libraryName  = 'pptxgen';

if (env === 'build') {
    plugins.push(new UglifyJsPlugin({
        minimize: true
    }));
    outputFile = libraryName  + '.min.js';
} else {
    outputFile = libraryName  + '.js';
}

var config = {
    entry: __dirname + '/main.js',
    devtool: 'source-map',
    output: {
        path: __dirname + '/lib',
        filename: outputFile,
        library: 'PptxGenJS' ,
        libraryTarget: 'var',
        umdNamedDefine: true,
        external: {
            jquery: 'jQuery',
            $: 'jQuery',
            Zip: 'JSZip',
            gObjPptxShapes: 'gObjPptxShapes',
            gObjPptxMasters: 'gObjPptxMasters'
        }
    },
    module: {
        loaders: [{
            test: /\.js$/,
            loader: "babel",
            exclude: /node_modules/
        }]
    },
    resolve: {
        root: path.resolve('./src'),
        extensions: ["", ".js"]
    },
    plugin: plugins
}

module.exports = config;
