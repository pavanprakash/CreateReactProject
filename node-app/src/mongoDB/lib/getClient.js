const MongoClient = require("mongodb").MongoClient;
const fs = require("fs");
const environment = process.env.NODE_ENV || "test";
const config = require("../../config/index");

const {
  MONGO_ADMIN_TEST_PASSWORD,
  MONGO_ADMIN_PROD_PASSWORD,
  MONGO_SSL_PASSWORD
} = process.env;

// console.log(__dirname+"/../")

let connectionstring, ca, cert, key;
switch (environment) {
  case "development":
    connectionstring = `mongodb://mongo_admin:${MONGO_ADMIN_TEST_PASSWORD}@londattst01,londattst02,londattst03/RMessageBus-Workflow?replicaSet=test&authSource=admin`;
    ca = fs.readFileSync(__dirname + "/../certificates/local/cacert.pem");
    cert = fs.readFileSync(__dirname + `/../certificates/local/cert.pem`);
    key = fs.readFileSync(__dirname + `/../certificates/local/key.pem`);
    break;
  case "production":
    connectionstring = config().mongoConnectionString;
    ca = fs.readFileSync(__dirname + "/../certificates/cacert.pem");
    cert = fs.readFileSync(__dirname + `/../certificates/cert.pem`);
    key = fs.readFileSync(__dirname + `/../certificates/key.pem`);
    break;
  default:
    connectionstring = config().mongoConnectionString;
    ca = fs.readFileSync(__dirname + "/../certificates/cacert.pem");
    cert = fs.readFileSync(__dirname + `/../certificates/cert.pem`);
    key = fs.readFileSync(__dirname + `/../certificates/key.pem`);
    break;
}

function getClient() {
  let connectOptions;

  connectOptions = {
    sslCA: ca,
    sslKey: key,
    sslCert: cert,
    useNewUrlParser: true
  };
  return new Promise((resolve, reject) => {
    MongoClient.connect(connectionstring, connectOptions, function(
      err,
      client
    ) {
      if (err) {
        reject(err);
      } else {
        resolve(client);
      }
    });
  });
}

module.exports = getClient;
