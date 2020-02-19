let config;

function getConfig(env = process.env.NODE_ENV || "production") {
  console.log(env);
  switch (env) {
    case "development":
      config = require("./development");
      break;
    case "production":
      config = require("./production");
      break;
    default:
      config = require("./production");
      break;
  }
  return config;
}

module.exports = getConfig;
