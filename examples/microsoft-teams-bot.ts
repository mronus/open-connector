import {
  microsoftTeamsBotActionHandlers,
  validateTeamsBotCredential,
} from "../src/providers/microsoft_teams_bot/runtime.ts";
/**
 * Mint a Microsoft Teams bot token and optionally send a message.
 *
 * Register an Azure Bot, create a client secret, then run:
 *
 *   TEAMS_BOT_APP_ID=... \
 *   TEAMS_BOT_CLIENT_SECRET=... \
 *   TEAMS_BOT_TENANT_ID=... \
 *   TEAMS_BOT_SERVICE_URL=https://smba.trafficmanager.net/teams/ \
 *   TEAMS_BOT_CONVERSATION_ID=... \
 *   node examples/microsoft-teams-bot.ts
 *
 * The service URL and conversation ID come from an inbound activity the bot
 * received. Without them the example only validates the credentials.
 */
import { createProviderFetch } from "../src/providers/provider-runtime.ts";

const appId = process.env.TEAMS_BOT_APP_ID?.trim();
const clientSecret = process.env.TEAMS_BOT_CLIENT_SECRET?.trim();
const tenantId = process.env.TEAMS_BOT_TENANT_ID?.trim();
const serviceUrl = process.env.TEAMS_BOT_SERVICE_URL?.trim();
const conversationId = process.env.TEAMS_BOT_CONVERSATION_ID?.trim();

if (!appId || !clientSecret || !tenantId) {
  console.log(
    "Skip Microsoft Teams Bot example: set TEAMS_BOT_APP_ID, TEAMS_BOT_CLIENT_SECRET, and TEAMS_BOT_TENANT_ID.",
  );
} else {
  const fetcher = createProviderFetch();
  const validation = await validateTeamsBotCredential({ appId, clientSecret, tenantId }, fetcher);
  console.log(`Validated bot ${validation.profile?.accountId}`);

  if (serviceUrl && conversationId) {
    const result = await microsoftTeamsBotActionHandlers.send_message(
      { serviceUrl, conversationId, text: "Hello from OpenConnector" },
      { appId, clientSecret, tenantId, fetcher },
    );
    console.log("Sent message:", result);
  } else {
    console.log("Set TEAMS_BOT_SERVICE_URL and TEAMS_BOT_CONVERSATION_ID to also send a message.");
  }
}
