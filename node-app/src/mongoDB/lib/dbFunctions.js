const getClient = require("./getClient");
const verifyMongoDBHostNames = require("./utils");

getRecords = async (dbName, collectionName, queryObject) => {
  const client = await getClient();
  // const hosts = client.s.options.servers.map(d => d.host);
  // verifyMongoDBHostNames(hosts, process.env.NODE_ENV);
  const db = client.db(dbName);

  const collection = db.collection(collectionName);
  try {
    const cursor = await collection.find(queryObject).sort({ _id: -1 });
    let records = [];
    for (
      let data = await cursor.next();
      data != null;
      data = await cursor.next()
    ) {
      records.push(data);
    }
    if (records.length == 0) {
      records = null;
      console.log("unable to fetch records for queryobject");
      // throw new Error(`unable to fetch records for queryobject`);
    }
    return records;
  } catch (error) {
    console.log(error);
  }
};
insertRecords = async (dbName, collectionName, testData) => {
  const client = await getClient();
  // const hosts = client.s.options.servers.map(d => d.host);
  // verifyMongoDBHostNames(hosts, process.env.NODE_ENV);
  const db = client.db(dbName);
  const collection = db.collection(collectionName);
  let result;
  try {
    result = await collection.insert(testData);
  } catch (err) {
    console.log(err);
  }
  if (result == undefined) {
    console.error("Failed to insert records into mongo DB");
  }
};

module.exports = { getRecords, insertRecords };
