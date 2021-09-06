const devCerts = require("office-addin-dev-certs");
const { CleanWebpackPlugin } = require("clean-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const path = require("path");
require("dotenv").config();

module.exports = async (env, options) => {
    return {
        devtool: "source-map",
        entry: {
            autorunshared: "./src/runtime/autorunshared.js",
            autorunweb: "./src/runtime/autorunweb.js",
            taskpane: "./src/taskpane/taskpane.js",
        },
        resolve: {
            extensions: [".js"],
        },
        plugins: [
            new CleanWebpackPlugin(),
            new HtmlWebpackPlugin({
                filename: "autorunweb.html",
                template: "./src/runtime/autorunweb.html",
                chunks: [],
            }),
            new HtmlWebpackPlugin({
                filename: "taskpane.html",
                template: "./src/taskpane/taskpane.html",
                chunks: [],
            }),
            new CopyWebpackPlugin({
                patterns: [
                    {
                        to: "[name][ext]",
                        from: "manifest.xml",
                        transform(content) {
                            return content
                                .toString()
                                .replace(
                                    new RegExp("{hostname}", "g"),
                                    process.env.URL
                                );
                        },
                    },
                    {
                        from: "./assets",
                        to: "assets",
                        force: true,
                    },
                ],
            }),
        ],
        devServer: {
            headers: {
                "Access-Control-Allow-Origin": "*",
            },
            https:
                options.https !== undefined
                    ? options.https
                    : await devCerts.getHttpsServerOptions().then((config) => {
                          // Unsuported key.
                          delete config.ca;
                          return config;
                      }),
            port: process.env.npm_package_config_dev_server_port || 3000,
        },
    };
};
