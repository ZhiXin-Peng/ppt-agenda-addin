/* webpack.config.js */
const path = require('path');
const HtmlWebpackPlugin = require('html-webpack-plugin');
const CopyWebpackPlugin = require('copy-webpack-plugin');
const webpack = require('webpack');
const devCerts = require('office-addin-dev-certs');

/**
 * 关键改动：
 * - 使用 office-addin-dev-certs.getHttpsServerOptions('localhost') 拿到已信任的证书
 * - 把 key/cert/ca 传给 webpack-dev-server 的 server.options
 * - 仍然使用 https://localhost:3000 与 manifest.xml 保持一致
 */
module.exports = async () => {
  // 拿到“已被系统信任”的 localhost 证书三件套
  const httpsOptions = await devCerts.getHttpsServerOptions('localhost');

  return {
    mode: 'development',
    entry: './src/taskpane/taskpane.ts',
    output: {
      path: path.resolve(__dirname, 'dist'),
      filename: 'taskpane.js',
      clean: true,
    },
    devtool: 'inline-source-map',
    resolve: {
      extensions: ['.ts', '.tsx', '.js'],
      // 禁用 Node 内置模块（避免某些库间接触发 node:* 解析）
      fallback: {
        fs: false, https: false, http: false, crypto: false, stream: false,
        path: false, url: false, buffer: false, zlib: false, util: false,
      },
      alias: {
        'node:fs': false, 'node:https': false, 'node:http': false, 'node:crypto': false,
        'node:stream': false, 'node:path': false, 'node:url': false, 'node:buffer': false,
        'node:zlib': false, 'node:util': false,
      },
    },
    module: {
      rules: [
        { test: /\.tsx?$/, use: 'ts-loader', exclude: /node_modules/ },
      ],
    },
    plugins: [
      new HtmlWebpackPlugin({
        template: './src/taskpane/taskpane.html',
        filename: 'taskpane.html',
        inject: 'body',
      }),
      new CopyWebpackPlugin({
        patterns: [{ from: 'assets', to: 'assets', noErrorOnMissing: true }],
      }),
      // 彻底忽略对 node:* 的尝试解析
      new webpack.IgnorePlugin({
        resourceRegExp: /^node:(fs|https|http|crypto|stream|path|url|buffer|zlib|util)$/,
      }),
    ],
    devServer: {
      static: { directory: path.join(__dirname, 'dist'), publicPath: '/' },
      port: 3000,
      server: {
        type: 'https',
        options: {
          key: httpsOptions.key,
          cert: httpsOptions.cert,
          ca: httpsOptions.ca,
        },
      },
      allowedHosts: 'all',
      headers: { 'Access-Control-Allow-Origin': '*' },
      hot: true,
      open: false,
    },
    performance: { hints: false },
    stats: { errorDetails: true, children: true },
  };
};
