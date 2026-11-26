/** Resolve historical authors conservatively; never guess between matching names. Inactive users still count towards ambiguity. */
/** @template {{name:string,active?:boolean}} T @param {T[]} users @param {string} name @returns {T|undefined} */
export function resolveHistoricalReplyAuthor(users,name) {
  const normalized = name.trim().toLowerCase();
  if (!normalized) return undefined;
  const exact = users.filter(u=>u.name.trim().toLowerCase()===normalized);
  const matches = exact.length ? exact : users.filter(u=>u.name.trim().toLowerCase().split(/\s+/)[0]===normalized);
  return matches.length===1 && matches[0].active!==false ? matches[0] : undefined;
}

/** @param {{active?:boolean}} user */
export function isActiveReplyRecipient(user) { return user.active!==false; }
