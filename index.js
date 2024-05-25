require('dotenv').config();
const express = require('express');
const session = require('express-session');

require('isomorphic-fetch');
const {PublicClientApplication, ConfidentialClientApplication} = require('@azure/msal-node');
const {Client} = require('@microsoft/microsoft-graph-client');

const app = express();

const clientId = process.env.AZURE_AD_CLIENT_ID;
const clientSecret = process.env.AZURE_AD_CLIENT_SECRET;
const tenantId = process.env.AZURE_AD_TENANT_ID;
const scopes = ["https://graph.microsoft.com/.default"];

// Configure session
app.use(session({
    secret: 'your-secret-key',
    resave: false,
    saveUninitialized: true,
}));

// Debug: Print environment variables
console.log('Client ID:', clientId);
console.log('Tenant ID:', tenantId);
console.log('Client Secret:', clientSecret);

// MSAL configuration

const msalConfig = {
    auth: {
        clientId,
        authority: `https://login.microsoftonline.com/${tenantId}`,
        clientSecret // Ensure this is set
    },
    system: {
        loggerOptions: {
            loggerCallback(loglevel, message) {
                console.log(message);
            },
            piiLoggingEnabled: false,
            logLevel: 'Info'
        }
    }
};

const pca = new PublicClientApplication(msalConfig);

const ccaConfig = {
    auth: {
        clientId,
        authority: `https://login.microsoftonline.com/${tenantId}`,
        clientSecret
    }
};

const cca = new ConfidentialClientApplication(ccaConfig)

// Redirect URL
const REDIRECT_URI = 'http://localhost:3000/azure-callback';

// Routes
app.get('/', (req, res) => {
    res.send('<a href="/auth">Login with Microsoft</a>');
});

app.get('/auth', (req, res) => {
    const authCodeUrlParameters = {
        scopes,
        redirectUri: REDIRECT_URI,
    };

    return pca.getAuthCodeUrl(authCodeUrlParameters)
        .then((response) => {
            res.redirect(response);
        })
        .catch((error) => console.log(JSON.stringify(error)));
});

app.get('/azure-callback', (req, res) => {
    let tokenRequest = {
        code: req.query.code,
        scopes: scopes,
        redirectUri: REDIRECT_URI,
    };

    pca.acquireTokenByCode(tokenRequest)
        .then((response) => {
            req.session.accessToken = response.accessToken;
            tokenRequest = {scopes}
            return cca.acquireTokenByClientCredential(tokenRequest)
        })
        .then((response) => {
            req.session.clientAccessToken = response.accessToken;
            res.redirect('/emails');
        })
        .catch((error) => {
            console.error(error);
            res.status(500).send(error);
        });
});

app.get('/emails', (req, res) => {
    if (!req.session.accessToken) {
        return res.redirect('/');
    }
    if (!req.session.clientAccessToken) {
        return res.redirect('/');
    }

    const client = Client.init({
        authProvider: (done) => {
            done(null, req.session.accessToken);
        },
    });

    return client.api('/me/messages').get()
        .then(messages => res.json(messages))
        .catch(err => {
            console.error(err)
            res.status(500).send(err)
        })
});

app.listen(3000, () => {
    console.log('Server started on http://localhost:3000');
});