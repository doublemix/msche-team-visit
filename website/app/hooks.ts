"use client";

import { useSearchParams } from "next/navigation";
import { getNow } from "./utils";

export function useNow() {
  let searchParams = useSearchParams();
  let nowFromSearch = searchParams.get("now");
  return getNow(nowFromSearch);
}
