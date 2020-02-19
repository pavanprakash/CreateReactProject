function verifyMongoDBHostNames(hosts, env) {
  const { intEnvironmentMongoDBHosts: int, testEnvironmentMongoDBHosts: test } =
    config.default || config;
  const expected = {
    test,
    int
  }[env];

  if (
    hosts[0] !== expected[0] ||
    hosts[1] !== expected[1] ||
    hosts[2] !== expected[2]
  ) {
    console.log(chalk.green(`Expected mongodb hosts = ${expected}`));
    console.log(chalk.red(`Actual mongodb hosts = ${hosts}`));
    throw new Error("Unexpected host names when connecting to mongo db");
  } else {
    console.log(chalk.green(`MongoDB hosts = ${hosts}`));
  }
}

module.exports = verifyMongoDBHostNames;
