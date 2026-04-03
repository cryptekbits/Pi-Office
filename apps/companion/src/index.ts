import { CompanionServer } from "./server.js";

const server = new CompanionServer();
server.start().catch((error) => {
  console.error(error);
  process.exit(1);
});
