/*-----------------------------------------------------------------------------
A simple OAuthCard bot for the Microsoft Bot Framework. 
-----------------------------------------------------------------------------*/

var builder = require('botbuilder');
// var botbuilder_azure = require("botbuilder-azure");
// var https = require('https');

// Graph API SDK for Node
// var MicrosoftGraph = require("@microsoft/microsoft-graph-client");

var useEmulator = (process.env.NODE_ENV == 'development');
var useEmulator = true;
// setting up internal storage. Do not use in-proc storage for production!!!
var inMemoryStorage = new builder.MemoryBotStorage();

// Create chat connector for communicating with the Bot Framework Service
var connector = new builder.ChatConnector({
    appId: process.env["MicrosoftAppId"] || "fe8957be-3ae8-4c39-969c-00eadd544339",
    appPassword: process.env["MicrosoftAppPassword"] || "viprqAB_ssEADWS04910:^[",
    openIdMetadata: process.env['BotOpenIdMetadata']
});

var connectionName = process.env.CONNECTION_NAME || "NodeOAuthNGraphAAD";
var userEmail;

// container for MS graph client
var client;

// Unable to get this working properly, however this should be on the right track
//
// Since we are using Node, however, I think the SDK is probably the easiest way to go!
/*function getEmailFromGraph(token, callback) {
    console.log("making get request to Graph API");
    const options = {
        method: 'GET',
        Authorization: 'Bearer ' + token,
        url: "https://graph.microsoft.com/v1.0/me/",
        hostname: "graph.microsoft.com"
    }

    // See https://davidwalsh.name/nodejs-http-request
    var req = https.request(options, function(response) {
        response.setEncoding('utf8');
        var body = '';

        response.on('data', function(data) {
            body += data;
        });

        response.on('end', function() {
            try {
                //var parsed = JSON.parse(body);
                console.log("response: ", body);
            } catch(e) {
                console.log("Unable to parse request response: ", );
                callback(e);
            }
            callback(body);
        })
    })
    
    req.on('error', function(err) {
        console.log("HTTP GET error: " + err.message);
        callback(err);
    });

    // send request
    req.end();
    console.log("request sent.");
};*/

// Create your bot with a function to receive messages from the user
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
                // init Graph client with AAD token
                // client = MicrosoftGraph.Client.init({
                //     authProvider: function(done) {
                //         done(null, result.token); // token we've gotten from AAD v2
                //     }
                // });

                // Request from the graph the users email
                // after the promise resolves (then()), send the message to the user
                //
                // See https://github.com/microsoftgraph/msgraph-sdk-javascript for usage details
                // client.api('https://graph.microsoft.com/v1.0/me') // I wasn't able to successfully use just 'me', but the full URL works just fine!
                // .select("mail") // specifically selects the mail category from the information returned
                // .get() // execute get request
                // .then(function(res) {
                //     console.log("Response: ", res); // debugging purposes
                //                                     // used to see what the response looks like (it's a JavaScript object)
                //     userEmail = res["mail"];

                //     session.send('You are already signed in with token: ' + result.token + '\n\n'
                //             + 'Your email is: ' + userEmail);
                // }).catch(function(err) {
                //     console.log("Graph API GET Error: ", err);
                //     session.send('Uh oh! It looks like I can\'t communicate with the graph right now');
                // });

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
})
.set('storage', inMemoryStorage) // Register in memory storage
.on("event", (event) => {         // Handle 'event' activities
    if (event.name == 'tokens/response') {
        // received a TokenResponse, which is how the Azure Bot Service responds with the user token after an OAuthCard
        bot.loadSession(event.address, (err, session) => {
            var tokenResponse = event.value;
            session.send('You are now signed in with token: ' + tokenResponse.token);
            session.userData.activeSignIn = false;
        });
    }
});

connector.onInvoke((event, cb) => {
    if (event.name == 'signin/verifyState') {
        // received a MS Team's code verification Invoke Activity
        bot.loadSession(event.address, (err, session) => {
            var verificationCode = event.value.state;
            // Get the user token using the verification code sent by MS Teams
            connector.getUserToken(session.message.address, connectionName, verificationCode, (err, result) => {
                session.send('You are now signed in with token: ' + result.token);
                session.userData.activeSignIn = false;
                cb(undefined, {}, 200);
            });
        });
    } else {
        cb(undefined, {}, 200);
    }
});

if(useEmulator) {
    var restify = require('restify');
    // Setup Restify Server
    var server = restify.createServer();
    server.listen(process.env.port || process.env.PORT || 4000, function () {
        console.log('%s listening to %s', server.name, server.url); 
    });
    // Listen for messages from users 
    server.post('/api/messages', connector.listen());
} else {
    module.exports = connector.listen()
}