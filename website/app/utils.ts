export function singleOrDefault<T>(x: T[], defaultValue: T): T {
  if (x.length > 1) throw new Error("single found multiple");
  if (x.length === 1) return x[0];
  return defaultValue;
}

export function single<T>(x: T[]) {
  if (x.length > 1) throw new Error("single found multiple");
  if (x.length === 1) return x[0];
  throw new Error("single found no results");
}

export function _throw(error: unknown): never {
  throw error;
}

export function getNowFromSearch(searchParams: { now?: string }) {
  return getNow(searchParams?.now ?? null);
}

export function getNow(userNow: string | null) {
  let now: Date | null = null;
  if (userNow !== null) now = new Date(userNow);
  if (now !== null && !isFinite(now.valueOf())) now = null;
  return now?.valueOf() ?? Date.now();
}

type Group<T, K extends PropertyKey> = T[] & { key: K };
type Grouped<T, K extends PropertyKey> = Group<T, K>[];

export const groupBy = <T, K extends PropertyKey>(
  data: T[],
  keySelector: (item: T) => K,
): Grouped<T, K> => {
  let groupMap = new Map<K, Group<T, K>>();
  let groups: Group<T, K>[] = [];

  data.forEach((item) => {
    let key = keySelector(item);
    if (!groupMap.has(key)) {
      let group = Object.assign([], { key });
      groupMap.set(key, group);
      groups.push(group);
    }
    groupMap.get(key)!.push(item);
  });

  return groups;
};
