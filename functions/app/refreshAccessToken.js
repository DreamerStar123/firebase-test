const axios = require('axios');
require("dotenv").config();

async function refreshGoogleAccessToken(refreshToken) {
  try {
    const response = await axios.post('https://oauth2.googleapis.com/token', {
      client_id: process.env.GCP_CLIENT_ID,
      client_secret: process.env.GCP_CLIENT_SECRET,
      refresh_token: refreshToken,
      grant_type: 'refresh_token',
    }, {
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
    });

    const { access_token, expires_in } = response.data;
    // console.log("New Access Token:", access_token);
    console.log("Expires In:", expires_in, "seconds");
    return access_token;
  } catch (error) {
    console.error("Error refreshing access token:", error.response?.data || error.message);
    throw error;
  }
}

async function refreshMicrosoftAccessToken(refreshToken) {
  try {
    const response = await axios.post('https://login.microsoftonline.com/common/oauth2/v2.0/token', {
      client_id: process.env.AZURE_CLIENT_ID,
      client_secret: process.env.AZURE_CLIENT_SECRET,
      refresh_token: refreshToken,
      grant_type: 'refresh_token',
    }, {
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
    });

    const { access_token, expires_in } = response.data;
    // console.log("New Access Token:", access_token);
    console.log("Expires In:", expires_in, "seconds");
    return access_token;
  } catch (error) {
    console.error("Error refreshing access token:", error.response?.data || error.message);
    throw error;
  }
}

module.exports = {
  refreshGoogleAccessToken,
  refreshMicrosoftAccessToken
}
