import { createCookieSessionStorage } from "react-router";
import { scryptSync, randomBytes, timingSafeEqual } from "crypto";

const { getSession, commitSession, destroySession } = createCookieSessionStorage({
  cookie: {
    name: "portal_session",
    httpOnly: true,
    maxAge: 60 * 60 * 24 * 7,
    path: "/portal",
    sameSite: "lax",
    secrets: [process.env.PORTAL_SESSION_SECRET || "portal-dev-secret-change-in-prod"],
    secure: process.env.NODE_ENV === "production",
  },
});

export { getSession, commitSession, destroySession };

export function hashPassword(password: string): string {
  const salt = randomBytes(16).toString("hex");
  const hash = scryptSync(password, salt, 64).toString("hex");
  return `${salt}:${hash}`;
}

export function verifyPassword(password: string, stored: string): boolean {
  try {
    const [salt, hash] = stored.split(":");
    const hashBuffer = Buffer.from(hash, "hex");
    const suppliedHash = scryptSync(password, salt, 64);
    return timingSafeEqual(hashBuffer, suppliedHash);
  } catch {
    return false;
  }
}
