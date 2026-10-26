import type { BotFrameworkSigningKey } from "./teams-bot-event-verification.ts";

import { generateKeyPair, SignJWT } from "jose";
import { describe, expect, it } from "vitest";
import { verifyTeamsBotEventRequest } from "./teams-bot-event-verification.ts";

const appId = "00000000-0000-0000-0000-000000000001";
const issuer = "https://api.botframework.com";
const serviceUrl = "https://smba.trafficmanager.net/teams/";
const now = 1_700_000_000_000;
const { privateKey, publicKey } = await generateKeyPair("RS256");
const otherKeys = await generateKeyPair("RS256");

interface TokenOptions {
  iss?: string;
  aud?: string;
  exp?: number | null;
  claimServiceUrl?: string;
  key?: CryptoKey;
}

async function token(options: TokenOptions = {}): Promise<string> {
  const jwt = new SignJWT({ serviceurl: options.claimServiceUrl ?? serviceUrl })
    .setProtectedHeader({ alg: "RS256", kid: "key-1" })
    .setIssuer(options.iss ?? issuer)
    .setAudience(options.aud ?? appId);
  if (options.exp !== null) jwt.setExpirationTime(options.exp ?? now / 1000 + 600);
  return jwt.sign(options.key ?? privateKey);
}

function activity(overrides: Record<string, unknown> = {}): Uint8Array {
  return new TextEncoder().encode(JSON.stringify({ type: "message", serviceUrl, channelId: "msteams", ...overrides }));
}

async function verify(
  tokenOptions: TokenOptions = {},
  body: Uint8Array = activity(),
  endorsements: string[] = ["msteams"],
) {
  const signingKey: BotFrameworkSigningKey = { key: publicKey, endorsements };
  return verifyTeamsBotEventRequest({
    authorization: `Bearer ${await token(tokenOptions)}`,
    rawBody: body,
    appId,
    now,
    resolveKey: async (_cloud, keyId) => (keyId === "key-1" ? signingKey : undefined),
  });
}

describe("Teams bot inbound activity verification", () => {
  it("accepts a token for the bot that matches the activity serviceUrl", async () => {
    await expect(verify()).resolves.toEqual({ valid: true, appId, serviceUrl, channelId: "msteams" });
  });

  it("rejects a token signed by another key", async () => {
    await expect(verify({ key: otherKeys.privateKey })).resolves.toEqual({ valid: false, reason: "token" });
  });

  it("rejects an unknown issuer and a token meant for another bot", async () => {
    await expect(verify({ iss: "https://attacker.example" })).resolves.toEqual({ valid: false, reason: "token" });
    await expect(verify({ aud: "other-bot" })).resolves.toEqual({ valid: false, reason: "app" });
  });

  it("rejects an expired token beyond the clock tolerance", async () => {
    await expect(verify({ exp: now / 1000 - 600 })).resolves.toEqual({ valid: false, reason: "timestamp" });
  });

  it("rejects a signed token that carries no expiry", async () => {
    await expect(verify({ exp: null })).resolves.toEqual({ valid: false, reason: "token" });
  });

  it("binds the token to the activity serviceUrl", async () => {
    await expect(verify({}, activity({ serviceUrl: "https://smba.trafficmanager.net/amer/" }))).resolves.toEqual({
      valid: false,
      reason: "service_url",
    });
  });

  it("requires the signing key to be endorsed for the activity channel", async () => {
    await expect(verify({}, activity(), ["webchat"])).resolves.toEqual({ valid: false, reason: "channel" });
    await expect(verify({}, activity({ channelId: undefined }), ["webchat"])).resolves.toMatchObject({ valid: true });
  });

  it("reports a malformed body only after the token is trusted", async () => {
    await expect(verify({}, new TextEncoder().encode("not json"))).resolves.toEqual({
      valid: false,
      reason: "payload",
    });
    await expect(
      verifyTeamsBotEventRequest({ authorization: "Bearer nope", rawBody: new Uint8Array(), appId, now }),
    ).resolves.toEqual({ valid: false, reason: "token" });
  });
});
