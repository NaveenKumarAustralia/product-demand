import bcrypt from "bcryptjs";
import prisma from "../db.server";
import { PORTAL_USERS_KEY } from "../portal-messages.server";
import type { OAuthClient, OAuthCode, OAuthStore, OAuthToken, PortalApprover, PortalUserDirectory } from "./oauth-core";

export const prismaOAuthStore: OAuthStore = {
  async createClient(client) {
    await prisma.mcpOAuthClient.create({ data: { id: client.id, name: client.name, redirectUris: client.redirectUris } });
  },
  async getClient(id) {
    const row = await prisma.mcpOAuthClient.findUnique({ where: { id } });
    if (!row) return null;
    return { id: row.id, name: row.name, redirectUris: Array.isArray(row.redirectUris) ? (row.redirectUris as unknown[]).map(String) : [] } satisfies OAuthClient;
  },
  async touchClient(id) {
    await prisma.mcpOAuthClient.update({ where: { id }, data: { lastUsedAt: new Date() } }).catch(() => {});
  },
  async saveCode(code) {
    await prisma.mcpOAuthCode.create({ data: code });
  },
  async claimCode(codeHash, now) {
    // updateMany with usedAt: null is atomic in Postgres: only one caller wins.
    const res = await prisma.mcpOAuthCode.updateMany({ where: { codeHash, usedAt: null }, data: { usedAt: now } });
    if (res.count !== 1) return null;
    return (await prisma.mcpOAuthCode.findUnique({ where: { codeHash } })) as OAuthCode | null;
  },
  async createToken(token) {
    await prisma.mcpOAuthToken.create({ data: token });
  },
  async findTokenByAccessHash(hash) {
    return (await prisma.mcpOAuthToken.findUnique({ where: { accessTokenHash: hash } })) as OAuthToken | null;
  },
  async findTokenByRefreshHash(hash) {
    return (await prisma.mcpOAuthToken.findUnique({ where: { refreshTokenHash: hash } })) as OAuthToken | null;
  },
  async markRotated(id, now) {
    const res = await prisma.mcpOAuthToken.updateMany({ where: { id, rotatedAt: null, revokedAt: null }, data: { rotatedAt: now } });
    return res.count === 1;
  },
  async revokeGrant(grantId, now) {
    await prisma.mcpOAuthToken.updateMany({ where: { grantId, revokedAt: null }, data: { revokedAt: now } });
  },
  async touchToken(id, now) {
    await prisma.mcpOAuthToken.update({ where: { id }, data: { lastUsedAt: now } }).catch(() => {});
  },
};

type RawPortalUser = {
  id?: unknown;
  name?: unknown;
  passwordHash?: unknown;
  role?: unknown;
  active?: unknown;
};

async function loadPortalUsers(): Promise<RawPortalUser[]> {
  const row = await prisma.portalSetting.findUnique({ where: { key: PORTAL_USERS_KEY }, select: { value: true } });
  return Array.isArray(row?.value) ? (row!.value as RawPortalUser[]) : [];
}

function toApprover(u: RawPortalUser): PortalApprover {
  const role = String(u.role ?? "user");
  return {
    id: String(u.id ?? ""),
    name: String(u.name ?? "").trim(),
    isAdmin: role === "superadmin" || role === "admin",
    active: u.active !== false,
  };
}

// A dummy hash so a login for an unknown user costs the same time as a real one.
const DUMMY_HASH = "$2a$10$CwTycUXWue0Thq9StjUM0uJ8pZsKXyJ2pJ1x0Kq3h5C9k1cX1uYyW";

export const portalUserDirectory: PortalUserDirectory = {
  async verifyLogin(username, password) {
    const users = await loadPortalUsers();
    const wanted = username.trim().toLowerCase();
    const user = users.find((u) => String(u.name ?? "").trim().toLowerCase() === wanted && u.active !== false);
    const hash = user && typeof user.passwordHash === "string" && user.passwordHash ? user.passwordHash : DUMMY_HASH;
    const ok = await bcrypt.compare(password, hash).catch(() => false);
    return ok && user ? toApprover(user) : null;
  },
  async getUser(id) {
    const users = await loadPortalUsers();
    const user = users.find((u) => String(u.id ?? "") === id);
    return user ? toApprover(user) : null;
  },
};
