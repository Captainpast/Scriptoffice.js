import * as path from "path"
import { fileURLToPath } from "url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

export default {
  entry: "./src/index.ts",
  //devtool: "inline-source-map",
  devtool: "source-map",
  module: {
    rules: [ { test: /\.ts$/, use: "ts-loader" }, ],
  },
  resolve: { extensions: [".ts", ".js"], },
  output: {
    path: path.resolve(__dirname, "dist"),
    filename: "script-office.min.js",
    library: "OfficeDocument",
    libraryTarget: "umd",
  },
  watchOptions: {
    ignored: "**/node_modules",
  },
};