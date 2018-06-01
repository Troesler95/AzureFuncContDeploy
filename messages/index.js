"use strict";
var builder = require("botbuilder");
var botbuilder_azure = require("botbuilder-azure");
var path = require('path');

var useEmulator = (process.env.NODE_ENV == 'development');

var connector = useEmulator ? new builder.ChatConnector() : new botbuilder_azure.BotServiceConnector({
    appId: process.env['MicrosoftAppId'],
    appPassword: process.env['MicrosoftAppPassword'],
    openIdMetadata: process.env['BotOpenIdMetadata']
});

// AAD Connection for this bot
var connectionName = process.env.CONNECTION_NAME || "NodeOAuthNGraphAAD";

/*----------------------------------------------------------------------------------------
* Bot Storage: This is a great spot to register the private state storage for your bot. 
* We provide adapters for Azure Table, CosmosDb, SQL Azure, or you can implement your own!
* For samples and documentation, see: https://github.com/Microsoft/BotBuilder-Azure
* ---------------------------------------------------------------------------------------- */

//var tableName = 'botdata';
//var azureTableClient = new botbuilder_azure.AzureTableClient(tableName, process.env['AzureWebJobsStorage']);
//var tableStorage = new botbuilder_azure.AzureBotStorage({ gzipData: false }, azureTableClient);

// for testing purposes exclusively. This should never be used in production!
var inMemoryStorage = new builder.MemoryBotStorage();

var bot = new builder.UniversalBot(connector, function (session) {
    if (session.message.text == 'signout') {
        // It is important to have a SignOut intent
        connector.signOutUser(session.message.address, connectionName,  (err, result) => {
            if (!err) {
                session.send('You are signed out.');
            } else {
                session.send('There was a problem signing you out.');                
            }
        });
    } else {
        // First check whether the Azure Bot Service already has a token for this user
        connector.getUserToken(session.message.address, connectionName, undefined, (err, result) => {
            if (result) {
                // If there is already a token, the bot can use it directly
                session.send('You are already signed in with token: ' + result.token);
            } else {
                // If there not is already a token, the bot can send an OAuthCard to have the user log in
                if (!session.userData.activeSignIn) {
                    session.send("Hello! Let's get you signed in!");
                    builder.OAuthCard.create(connector, session, connectionName, "Please sign in", "Sign in", (createSignInErr, signInMessage) =>
                    {
                        if (signInMessage) {
                            session.send(signInMessage);
                            session.userData.activeSignIn = true;
                        } else {
                            session.send("Something went wrong trying to sign you in.");
                        }     
                    });
                } else {
                    // Some clients require a 6 digit code validation so we can check that here
                    session.send("Let's see if that code works...");
                    connector.getUserToken(session.message.address, connectionName, session.message.text, (err2, tokenResponse) => {
                        if (tokenResponse) {
                            session.send('It worked! You are now signed in with token: ' + tokenResponse.token);
                            session.userData.activeSignIn = false;
                        } else {
                            session.send("Hmm, that code wasn't right");
                        }
                    });
                }
            }
        });
    }
}).on("event", (event) => {         // Handle 'event' activities
    if (event.name == 'tokens/response') {
        // received a TokenResponse, which is how the Azure Bot Service responds with the user token after an OAuthCard
        bot.loadSession(event.address, (err, session) => {
            let tokenResponse = event.value;
            session.send('You are now signed in with token: ' + tokenResponse.token);
            session.userData.activeSignIn = false;
        });
    }
});

bot.localePath(path.join(__dirname, './locale'));
bot.set('storage', inMemoryStorage);

bot.dialog('/', function (session) {
    session.send('You said ' + session.message.text);
});

if (useEmulator) {
    var restify = require('restify');
    var server = restify.createServer();
    server.listen(3978, function() {
        console.log('test bot endpont at http://localhost:3978/api/messages');
    });
    server.post('/api/messages', connector.listen());    
} else {
    module.exports = connector.listen();
}
