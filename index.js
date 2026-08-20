require('dotenv').config();
const fs = require('fs');
const path = require('path');
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

        messages.forEach((msg, index) => {
            const senderEmail = msg.from?.emailAddress?.address;

            if (checkBlacklist(senderEmail)) {
                client.api(`/me/messages/${msg.id}`).delete().then(() => {
                    console.log(`Message from ${senderEmail} deleted successfully.`);
                }).catch((deleteError) => {
                    console.error(`Failed to delete message from ${senderEmail}:`, deleteError);
                });
            }
        });

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