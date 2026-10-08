/* eslint-disable no-undef */

const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const webpack = require("webpack");

const urlDev = "https://localhost:3000/";
// Where the production add-in is served: the customer's PlaceOS domain under /outlook-addin/ (one manifest per customer).
// Override per customer, e.g. ADDIN_URL=https://<customer>.placeos.com/outlook-addin/ npm run build
const urlProd = process.env.ADDIN_URL || "https://placeos-dev.aca.im/outlook-addin/";

// PlaceOS domain that auth.cr and Staff API requests are proxied to in local development.
const placeosDevDomain = "https://placeos-dev.aca.im";

/** The S3 host in a dev map proxy path (/__map-proxy/<host>/...), or null. */
function mapProxyHost(path) {
  const match = /^\/__map-proxy\/([^/]+)\//.exec(path || "");
  return match && /(^|\.)amazonaws\.com$/i.test(match[1]) ? match[1] : null;
}

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const config = {
    devtool: "source-map",
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      taskpane: ["./src/taskpane/taskpane.ts", "./src/taskpane/taskpane.html"],
      dialog: ["./src/taskpane/fallback/fallbackauthdialog.ts"],
      // Personal tab for the Teams / Outlook / Microsoft 365 app bar (see app-package/).
      app: ["./src/taskpane/app.ts"],
    },
    output: {
      clean: true,
      // PlaceOS serves the add-in without Cache-Control headers, so webviews cache fixed names like taskpane.js
      // heuristically and keep running old code after a deploy. Hashed names make each build's HTML load new bundles.
      filename: dev ? "[name].js" : "[name].[contenthash].js",
    },
    resolve: {
      extensions: [".ts", ".html", ".js"],
    },
    module: {
      rules: [
        {
          test: /\.ts$/,
          exclude: /node_modules/,
          use: {
            loader: "babel-loader",
            options: {
              presets: ["@babel/preset-typescript"],
            },
          },
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          resourceQuery: { not: [/app/] },
          use: "html-loader",
        },
        {
          // app.html reuses the task pane template without office.js, which isn't needed outside
          // Office and interferes with the history API that TeamsJS hosts rely on.
          test: /\.html$/,
          exclude: /node_modules/,
          resourceQuery: /app/,
          use: {
            loader: "html-loader",
            options: {
              preprocessor: (content) =>
                content.replace(/<script[^>]*office\.js[^>]*><\/script>/, ""),
            },
          },
        },
        {
          test: /\.(png|jpg|jpeg|gif|ico)$/,
          type: "asset/resource",
          generator: {
            filename: "assets/[name][ext][query]",
          },
        },
      ],
    },
    plugins: [
      new webpack.DefinePlugin({
        // Dev-only tooling, e.g. revealing the raw Entra token. Dropped from production builds.
        __DEV_TOOLS__: JSON.stringify(dev),
      }),
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane"],
      }),
      new HtmlWebpackPlugin({
        filename: "app.html",
        template: "./src/taskpane/taskpane.html?app",
        chunks: ["polyfill", "app"],
      }),
      new HtmlWebpackPlugin({
        filename: "auth.html",
        template: "./src/taskpane/fallback/auth.html",
        chunks: [],
      }),
      new HtmlWebpackPlugin({
        filename: "dialog.html",
        template: "./src/taskpane/fallback/dialog.html",
        chunks: ["dialog"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "assets/*",
            to: "assets/[name][ext][query]",
          },
          {
            from: "manifest*.xml",
            to: "[name]" + "[ext]",
            transform(content) {
              if (dev) {
                return content;
              } else {
                return content.toString().replace(new RegExp(urlDev, "g"), urlProd);
              }
            },
          },
        ],
      }),
    ],
    devServer: {
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      server: {
        type: "https",
        options:
          env.WEBPACK_BUILD || options.https !== undefined
            ? options.https
            : await getHttpsOptions(),
      },
      port: process.env.npm_package_config_dev_server_port || 3000,
      // Keep PlaceOS calls same-origin, matching production hosting on the PlaceOS domain.
      // Trailing slashes matter: "/auth" would also match the MSAL redirect page "/auth.html".
      proxy: [
        {
          context: ["/auth/", "/api/"],
          target: placeosDevDomain,
          changeOrigin: true,
          secure: true,
        },
        {
          // Floor plans (level map_id) are on S3, whose CORS rules allow the customer domain but not localhost.
          // floor-map.ts sends them here in development as /__map-proxy/<s3 host>/<path>. S3 hosts only.
          context: (pathname) => mapProxyHost(pathname) !== null,
          target: "https://s3.amazonaws.com",
          router: (req) => `https://${mapProxyHost(req.url)}`,
          pathRewrite: (path) => path.replace(/^\/__map-proxy\/[^/]+/, ""),
          changeOrigin: true,
          secure: true,
        },
      ],
    },
  };

  return config;
};
