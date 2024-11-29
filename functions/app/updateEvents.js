const admin = require('firebase-admin');
const { google } = require('googleapis');
const { Client } = require('@microsoft/microsoft-graph-client');
const serviceAccount = require('./firebase-adminsdk.json'); // Path to your service account JSON file
const { refreshGoogleAccessToken, refreshMicrosoftAccessToken } = require('./refreshAccessToken');
require("dotenv").config();

admin.initializeApp({
    credential: admin.credential.cert(serviceAccount),
});

const db = admin.firestore();

const getAllUsers = async () => {
    try {
        const listUsersResult = await admin.auth().listUsers();
        return listUsersResult.users;
    } catch (error) {
        console.error("Error listing users:", error);
        return null;
    }
}

const getUsers = async () => {
    try {
        let nextPageToken;
        const googleUsers = [];
        const microsoftUsers = [];

        do {
            const listUsersResult = await admin.auth().listUsers(1000, nextPageToken); // Batch of 1000 users
            listUsersResult.users.forEach((userRecord) => {
                // Check if user signed in with Google
                const isGoogleUser = userRecord.providerData.some(
                    (provider) => provider.providerId === "google.com"
                );
                const isMicrosoftUser = userRecord.providerData.some(
                    (provider) => provider.providerId === "microsoft.com"
                );
                if (isGoogleUser)
                    googleUsers.push(userRecord);
                if (isMicrosoftUser)
                    microsoftUsers.push(userRecord);
            });
            nextPageToken = listUsersResult.pageToken;
        } while (nextPageToken);

        console.log("Google Users:", googleUsers.length);
        console.log("Microsoft Users:", microsoftUsers.length);
        return { googleUsers, microsoftUsers };
    } catch (error) {
        console.error("Error listing users:", error);
        return null;
    }
}

const getGoogleCalendars = async (accessToken) => {
    const oauth2Client = new google.auth.OAuth2();
    oauth2Client.setCredentials({ access_token: accessToken });

    const calendar = google.calendar({ version: 'v3', auth: oauth2Client });
    const res = await calendar.calendarList.list();
    const calendars = res.data;

    return calendars;
}

const getGoogleEvents = async (accessToken) => {
    const oauth2Client = new google.auth.OAuth2();
    oauth2Client.setCredentials({ access_token: accessToken });

    const calendar = google.calendar({ version: 'v3', auth: oauth2Client });

    const now = new Date();
    const timeMin = now.toISOString();
    const timeMax = new Date(now.getTime() + 7 * 24 * 60 * 60 * 1000).toISOString(); // 7 days from now
    const res = await calendar.events.list({
        calendarId: 'primary',
        timeMin,
        timeMax,
        singleEvents: true,
        orderBy: 'startTime',
    });
    const events = res.data;

    return events;
}

const getMicrosoftEvents = async (accessToken) => {
    let events = [];
    const client = Client.init({
        authProvider: (done) => {
            done(null, accessToken);
        }
    });

    try {
        const now = new Date();
        const timeMin = now.toISOString();
        const timeMax = new Date(now.getTime() + 7 * 24 * 60 * 60 * 1000).toISOString(); // 7 days from now
        events = await client.api(`/me/calendarview?startDateTime=${timeMin}&endDateTime=${timeMax}`).get();
    } catch (error) {
        console.error(`Error fetching calendar events: ${error}`);
    }
    return events;
}

const updateEvents = async () => {
    try {
        const users = await getAllUsers();
        if (users) {
            for (let index = 0; index < users.length; index++) {
                const user = users[index];
                const docRef = db.collection("users").doc(user.uid);
                const userData = (await docRef.get()).data();

                console.log(`  ${userData.email}`);

                // Google calendar
                if (userData.google) {
                    const accessToken = await refreshGoogleAccessToken(userData.google);
                    const events = await getGoogleEvents(accessToken);
                    // Save the document
                    await docRef.update({ google_calendars: events.items });
                    console.log(`Google calendar updated for ${userData.email}`);
                }

                // Outlook calendar
                if (userData.microsoft) {
                    const accessToken = await refreshMicrosoftAccessToken(userData.microsoft);
                    const events = await getMicrosoftEvents(accessToken);
                    // Save the document
                    await docRef.update({ outlook_calendars: events.value });
                    console.log(`Outlook calendar updated for ${userData.email}`);
                }
            }
        }
    } catch (e) {
        console.error("Error adding document: ", e.status);
    }
}

module.exports = updateEvents;
