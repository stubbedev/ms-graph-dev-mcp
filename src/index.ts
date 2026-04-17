#!/usr/bin/env node
import { GraphMcpServer } from "./server.js";

async function main(): Promise<void> {
  const server = new GraphMcpServer();
  await server.start();
}

main().catch((err) => {
  console.error("Fatal error:", err);
  process.exit(1);
});
