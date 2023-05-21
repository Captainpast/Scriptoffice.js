import path from "path"
import { fileURLToPath } from "url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

export default (env, argv) => { return {
  entry: "./src/index.ts",
  devtool: argv.mode == "development"? "source-map" : undefined,
  module: {
    rules: [ { test: /\.ts$/, use: "ts-loader" }, ],
  },
  resolve: { extensions: [".ts", ".js"], },
  output: {
    path: path.resolve(__dirname, "dist"),
    filename: "script-office.min.js",
    library: "OfficeDocument",
    libraryTarget: "umd",
    globalObject: "globalThis",
    clean: true
  },
  watchOptions: {
    ignored: "**/node_modules",
  }
}};