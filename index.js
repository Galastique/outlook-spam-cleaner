const path = require('path');
require('dotenv').config({ path: path.join(__dirname, '.env') });
const fs = require('fs');
const msal = require('@azure/msal-node');
const { Client } = require('@microsoft/microsoft-graph-client');
require('isomorphic-fetch');

const blocklistPath = path.join(__dirname, 'blocked_senders.txt');
const tokenCachePath = path.join(__dirname, '.msal-cache.json');
const blocklist = fs
    .readFileSync(blocklistPath, 'utf8')
    .split(/\r?\n/)
    .map((line) => line.trim().toLowerCase())
    .filter(Boolean);

const msalConfig = {
    auth: {
        clientId: process.env.CLIENT_ID,
        authority: 'https://login.microsoftonline.com/consumers',
    },
};

const pca = new msal.PublicClientApplication(msalConfig);

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

async function loadTokenCache() {
    if (!fs.existsSync(tokenCachePath)) {
        return;
    }

    const cacheData = fs.readFileSync(tokenCachePath, 'utf8');
    if (cacheData) {
        pca.getTokenCache().deserialize(cacheData);
    }
}

function saveTokenCache() {
    const serializedCache = pca.getTokenCache().serialize();
    fs.writeFileSync(tokenCachePath, serializedCache, 'utf8');
}

async function getAccessToken() {
    await loadTokenCache();

    const scopes = ['Mail.ReadWrite', 'offline_access', 'openid', 'profile'];
    const tokenCache = pca.getTokenCache();
    const accounts = await tokenCache.getAllAccounts();

    if (accounts.length > 0) {
        try {
            const silentResponse = await pca.acquireTokenSilent({
                account: accounts[0],
                scopes,
            });

            saveTokenCache();
            return silentResponse.accessToken;
        } catch (silentError) {
            console.log('Cached token unavailable or expired. Falling back to interactive sign-in.');
        }
    }

    const interactiveRequest = {
        scopes,
        openBrowser: async (url) => {
            console.log(`Open this URL in your browser to sign in:\n${url}\n`);
        },
    };

    const response = await pca.acquireTokenInteractive(interactiveRequest);
    saveTokenCache();
    return response.accessToken;
}

async function deleteMessageWithRetry(client, messageId, senderEmail, retries = 5) {
    for (let attempt = 1; attempt <= retries; attempt++) {
        try {
            await client.api(`/me/messages/${messageId}`).delete();
            console.log(`Message from ${senderEmail} deleted successfully.`);
            return;
        } catch (error) {
            if (attempt === retries) {
                console.error(`Failed to delete message from ${senderEmail} after ${retries} attempts:`, error);
            } else {
                await sleep(1000);
            }
        }
    }
}

async function runSpamCleaner() {
    try {
        const accessToken = await getAccessToken();

        const client = Client.init({
            authProvider: (done) => done(null, accessToken),
        });

        // Fetch top 10 messages from spam folder
        const response = await client
            .api('/me/mailFolders/junkEmail/messages')
            .top(10)
            .select('id,from')
            .get();

        const messages = response.value;

        if (!messages || messages.length === 0) {
            return;
        }

        for (const msg of messages) {
            const senderEmail = msg.from?.emailAddress?.address;

            if (checkBlacklist(senderEmail.toLowerCase())) {
                await deleteMessageWithRetry(client, msg.id, senderEmail);
                await sleep(250); // Pause briefly between deletions to avoid network socket drops
            }
        }

    } catch (error) {
        console.error('Graph API Error:', error);
    }
}

function checkBlacklist(email) {
    if (!email || typeof email !== 'string') {
        return false;
    }

    return blocklist.some((blockedEntry) => {
        return email === blockedEntry || email.endsWith(`@${blockedEntry}`);
    });
}

runSpamCleaner();
setInterval(runSpamCleaner, 60 * 60 * 1000);