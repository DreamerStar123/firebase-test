require("dotenv").config();
const express = require("express");
const bodyParser = require("body-parser");
const axios = require("axios");
const cors = require('cors');
const logger = require("firebase-functions/logger");

// OAuth URLs
const AUTHORIZE_URL = `https://login.microsoftonline.com/${process.env.AZURE_TENANT_ID}/oauth2/v2.0/authorize`;
const TOKEN_URL = `https://login.microsoftonline.com/${process.env.AZURE_TENANT_ID}/oauth2/v2.0/token`;
let REDIRECT_URI;
let FRONT_ORIGIN;

const NODE_ENV = 'development';
if (NODE_ENV == 'production') {
  REDIRECT_URI = 'https://msal./auth/callback';
  FRONT_ORIGIN = 'https://turbotabs.com';
} else {
  REDIRECT_URI = 'http://localhost:5001/my-fire-test-1/us-central1/msal/auth/callback';
  FRONT_ORIGIN = 'http://localhost:5173';
}

const app = express();
app.use(bodyParser.urlencoded({ extended: true }));
app.use(cors({ origin: FRONT_ORIGIN }));

// Step 1: Redirect user to Microsoft's authorization page
app.get("/auth", (req, res) => {
  const params = new URLSearchParams({
    client_id: process.env.AZURE_CLIENT_ID,
    response_type: "code",
    redirect_uri: REDIRECT_URI,
    scope: "Calendars.ReadWrite offline_access",
    response_mode: "query",
    // state: "12345", // Optional for CSRF protection
  });

  res.redirect(`${AUTHORIZE_URL}?${params.toString()}`);
});

// Step 2: Handle callback and exchange authorization code for tokens
app.get("/auth/callback", async (req, res) => {
  const { code, state, error } = req.query;

  if (error) {
    return res.send(`Error: ${error}`);
  }

  try {
    // Exchange authorization code for tokens
    const tokenResponse = await axios.post(
      TOKEN_URL,
      new URLSearchParams({
        client_id: process.env.AZURE_CLIENT_ID,
        client_secret: process.env.AZURE_CLIENT_SECRET,
        grant_type: "authorization_code",
        code: code,
        redirect_uri: REDIRECT_URI,
      }),
      {
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
      }
    );

    const tokens = tokenResponse.data;
    if (tokens.refresh_token) {
      const params = new URLSearchParams({
        res: 'success',
        sec: tokens.refresh_token,
      });
      res.redirect(`${FRONT_ORIGIN}/oauth2callback-ms?${params.toString()}`);
    } else
      res.redirect(`${FRONT_ORIGIN}/oauth2callback-ms?res=no-refresh-token`);
  } catch (err) {
    logger.error(err.response ? err.response.data : err.message);
    res.redirect(`${FRONT_ORIGIN}/oauth2callback-ms?res=error`);
  }
});

// Step 3: Use refresh token to get a new access token
app.post("/auth/refresh", async (req, res) => {
  const { refresh_token } = req.body;

  try {
    const tokenResponse = await axios.post(
      TOKEN_URL,
      new URLSearchParams({
        client_id: process.env.AZURE_CLIENT_ID,
        client_secret: process.env.AZURE_CLIENT_SECRET,
        grant_type: "refresh_token",
        refresh_token: refresh_token,
      }),
      {
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
      }
    );

    const tokens = tokenResponse.data;

    // Display the refreshed tokens
    res.json({
      access_token: tokens.access_token,
      refresh_token: tokens.refresh_token,
      expires_in: tokens.expires_in,
      scope: tokens.scope,
    });
  } catch (err) {
    logger.error(err.response ? err.response.data : err.message);
    res.status(500).send("Error refreshing token.");
  }
});

module.exports = app;
