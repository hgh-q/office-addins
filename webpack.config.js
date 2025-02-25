/* eslint-disable no-undef */
const path = require("path");
const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");

const urlDev = "https://localhost:3000/";
const urlProd = "https://www.contoso.com/"; // CHANGE THIS TO YOUR PRODUCTION DEPLOYMENT LOCATION

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
      commands: ['./src/commands/commands.js'],
      excel: ['./src/excel/main.js'],
      word: ['./src/word/main.js'],
      ppt: ['./src/ppt/main.js'],
      popup: "./src/dialogs/main.js"
    },
    output: {
      // 输出目录，构建后的文件将存放在 /build 下
      path: path.resolve(__dirname, 'build'),

      // 输出文件名
      filename: (pathData) => {
        // 根据不同的应用，指定不同的子目录
        if (pathData.chunk.name === 'excel') {
          return 'excel/[name].bundle.js'; // Excel 相关文件存放在 build/excel/
        } else if (pathData.chunk.name === 'word') {
          return 'word/[name].bundle.js'; // Word 相关文件存放在 build/word/
        } else if (pathData.chunk.name === 'powerpoint') {
          return 'powerpoint/[name].bundle.js'; // PowerPoint 相关文件存放在 build/powerpoint/
        }
        return '[name].bundle.js'; // 默认处理
      },
      clean: true, // 每次构建时清理输出目录
    },
    module: {
      rules: [
        {
          test: /\.(js|jsx)$/, // 支持 JSX
          exclude: /node_modules/,
          use: {
            loader: "babel-loader",
            options:
            {
              presets: ["@babel/preset-env", "@babel/preset-react"]
            },
          },
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader",
        },
        {
          test: /\.(png|jpg|jpeg|gif|ico)$/,
          type: "asset/resource",
          generator: {
            filename: "assets/[name][ext][query]",
          },
        },
        {
          test: /\.module\.css$/, // Match files ending with .module.css
          use: [
            'style-loader',       // Inject styles into DOM
            {
              loader: 'css-loader',
              options: {
                modules: true,     // Enable CSS Modules
              },
            },
          ],
        },
        {
          test: /\.css$/, // 匹配所有的 CSS 文件
          exclude: /\.module\.css$/, // 排除模块化 CSS 文件
          use: [
            'style-loader', // 将 CSS 加入到页面中
            'css-loader', // 解析 CSS
          ],
        },
      ],
    },
    plugins: [
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "assets/*", // 从源目录 assets 中复制所有文件
            to: "assets/[name][ext][query]", // 复制到输出目录的 assets 文件夹，并保留原有文件名和扩展名
          },
          {
            from: "src/manifest/manifest*.xml", // 匹配所有的 manifest*.xml 文件（例如 manifest1.xml、manifest2.xml）
            to: "[name][ext]", // 复制到输出目录，保留原文件名和扩展名
            transform(content) { // 对文件内容进行处理
              if (dev) {
                return content; // 如果是开发环境，直接返回原内容
              } else {
                // 如果是生产环境，将内容中的 dev URL 替换成生产环境的 URL
                return content.toString().replace(new RegExp(urlDev, "g"), urlProd);
              }
            },
          },
        ],
      }),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"],
      }), 
      new HtmlWebpackPlugin({
        filename: "popup.html",
        template: "./src/dialogs/popup.html",
        chunks: ["polyfill", "popup"]
      }),
      new HtmlWebpackPlugin({
        template: "./public/index.html",
        chunks: ["polyfill", "excel"],
        filename: 'excel.html',
      }),
      new HtmlWebpackPlugin({
        template: './public/index.html',
        chunks: ['word', 'polyfill'],
        filename: 'word.html',
      }),
      new HtmlWebpackPlugin({
        template: './public/index.html',
        chunks: ['ppt', 'polyfill'],
        filename: 'ppt.html',
      }),
    ],
    resolve: {
      alias: {
        '@': path.resolve(__dirname, 'src'),
        '@excel': path.resolve(__dirname, 'src/excel'),
        '@word': path.resolve(__dirname, 'src/word'),
        '@ppt': path.resolve(__dirname, 'src/ppt'),
      },
      extensions: [".html", ".js", '.jsx'],
    },
    devServer: {
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      server: {
        type: "https",
        options: env.WEBPACK_BUILD || options.https !== undefined ? options.https : await getHttpsOptions(),
      },
      // proxy: [
      // {
      //   context: ['/api'],
      //   target: 'http://127.0.0.1:5000',
      //   // secure: false,  // 允许不安全的证书（避免验证问题）
      // },
      // {
      //   context: ['/socket'],  // 如果你有其他 WebSocket 连接路径
      //   target: 'wss://127.0.0.1:5000',
      //   secure: false,         // 如果目标是 HTTPS/WSS 服务器
      //   ws: true,              // 开启 WebSocket 代理
      // },
      // ],
      port: process.env.npm_package_config_dev_server_port || 3000,
      hot: true, // 启用 React 组件热重载
      historyApiFallback: true, // 让 React Router 兼容
    },
  };

  return config;
};
