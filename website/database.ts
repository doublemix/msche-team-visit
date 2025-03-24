import { neon } from "@neondatabase/serverless";

export function getDatabase() {
  return neon(`${process.env.DATABASE_URL}`);
}
