/**
 * Import function triggers from their respective submodules:
 *
 * const {onCall} = require("firebase-functions/v2/https");
 * const {onDocumentWritten} = require("firebase-functions/v2/firestore");
 *
 * See a full list of supported triggers at https://firebase.google.com/docs/functions
 */

const {onRequest} = require("firebase-functions/v2/https");
const functions = require('firebase-functions');
const msalApp = require('./app/msal');
const updateEvents = require('./app/updateEvents');

// Create and deploy your first functions
// https://firebase.google.com/docs/functions/get-started

exports.msal = onRequest(msalApp);
exports.updateEvents = functions.pubsub.schedule('every 24 hours').onRun(updateEvents);
