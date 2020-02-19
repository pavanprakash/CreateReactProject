const proxy = require("http-proxy-middleware");
const config = require("./config/index");
const host = config().hostName;
const serverPort = config().serverPort;
const proxyUrl = `http://${host}:${serverPort}`;

module.exports = function(app) {
  app.use(
    "/",
    proxy({
      target: proxyUrl,
      changeOrigin: true
    })
  );
};
