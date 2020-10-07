// Copyright (c) Wictor Wilén. All rights reserved.
// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

const webpack = require('webpack');
const Dotenv = require('dotenv-webpack');
const ESLintPlugin = require('eslint-webpack-plugin');
const BundleAnalyzerPlugin = require('webpack-bundle-analyzer').BundleAnalyzerPlugin;
const TerserPlugin = require('terser-webpack-plugin');

const path = require('path');
const fs = require('fs');
const argv = require('yargs').argv;

const tsNameof = require("ts-nameof");

const debug = argv.debug !== undefined;
const lint = argv["linting"];

const nodeModules = {};
fs.readdirSync('node_modules')
    .filter(function (x) {
        return ['.bin'].indexOf(x) === -1;
    })
    .forEach(function (mod) {
        nodeModules[mod] = 'commonjs ' + mod;
    });

const config = [{
    entry: {
        server: [
            __dirname + '/src/app/server.ts'
        ],
    },
    mode: debug ? 'development' : 'production',
    output: {
        path: __dirname + '/dist',
        filename: '[name].js',
        devtoolModuleFilenameTemplate: debug ? '[absolute-resource-path]' : '[]'
    },
    externals: nodeModules,
    devtool: 'source-map',
    resolve: {
        extensions: [".ts", ".tsx", ".js"],
        alias: {}
    },
    target: 'node',
    node: {
        __dirname: false,
        __filename: false,
    },
    module: {
        rules: [{
            test: /\.tsx?$/,
            exclude: [/lib/, /dist/],
            loader: "ts-loader",
            options: {
                configFile: "tsconfig-server.json",
            }
        }]
    },
    plugins: []
},
{
    entry: {
        client: [
            __dirname + '/src/app/scripts/client.ts'
        ],
        authStart: [
            __dirname + '/src/app/scripts/auth/start.ts'
        ],
        authEnd: [
            __dirname + '/src/app/scripts/auth/end.ts'
        ]
    },
    mode: debug ? 'development' : 'production',
    output: {
        path: __dirname + '/dist/web/scripts',
        filename: '[name].js',
        chunkFilename: '[chunkhash].js',
        libraryTarget: 'umd',
        library: 'FindMsg',
        publicPath: '/scripts/'
    },
    externals: {},
    devtool: debug ? 'source-map' : 'nosources-source-map',
    resolve: {
        extensions: [".ts", ".tsx", ".js"],
        alias: {}
    },
    target: 'web',
    module: {
        rules: [{
            test: /\.tsx?$/,
            exclude: [/lib/, /dist/],
            loader: "ts-loader",
            options: {
                configFile: "tsconfig-client.json",
                getCustomTransformers: () => ({ before: [tsNameof] })
            }
        },
        {
            test: /\.(eot|svg|ttf|woff|woff2)$/,
            loader: 'file-loader?name=public/fonts/[name].[ext]'
        }
        ]
    },
    plugins: [
        new Dotenv({
            systemvars: true
        }),
        new BundleAnalyzerPlugin({
            analyzerMode: "static",
            defaultSizes: "gzip",
            openAnalyzer: false,
            reportFilename: "../../report.html"
        })
    ],
    performance: {
        maxEntrypointSize: 400000,
        maxAssetSize: 400000,
        assetFilter: function (assetFilename) {
            return assetFilename.endsWith('.js');
        }
    },
}
];

if (lint !== false) {
    // config[0].plugins.push(new ESLintPlugin({
    //     files: ['./src/app/*.ts']
    // }));
    config[1].plugins.push(new ESLintPlugin({
        files: ['./src/app/scripts/**/*.ts', './src/app/scripts/**/*.tsx']
    }));
}


module.exports = config;