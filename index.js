// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const dotenv = require('dotenv');
const path = require("path");
const restify = require('restify');
const teams = require('botbuilder-teams');
const fs = require('fs');
const { ConnectorClient, MicrosoftAppCredentials } = require('botframework-connector');

// Import required bot services.
// See https://aka.ms/bot-services to learn more about the different parts of a bot.
const { BotFrameworkAdapter, TurnContext } = require('botbuilder');

// This bot's main dialog.
const { TeamsBot } = require('./bot');

// Import required bot configuration
const ENV_FILE = path.join(__dirname, '.env');
dotenv.config({ path: ENV_FILE });

// Create HTTP server
const server = restify.createServer();
server.listen(process.env.port || process.env.PORT || 3978, () => {
    console.log(`\n${ server.name } listening to ${ server.url }`);
});

// Create adapter.
// See https://aka.ms/about-bot-adapter to learn more about .bot file its use and bot configuration.
const adapter = new BotFrameworkAdapter({
    appId: process.env.MicrosoftAppId,
    appPassword: process.env.MicrosoftAppPassword
});

adapter.use(new teams.TeamsMiddleware());

// Catch-all for errors.
adapter.onTurnError = async (context, error) => {
    // This check writes out errors to console log .vs. app insights.
    // NOTE: In production environment, you should consider logging this to Azure
    //       application insights.
    console.error(`\n [onTurnError] unhandled error: ${ error }`);

    // Send a trace activity, which will be displayed in Bot Framework Emulator
    await context.sendTraceActivity(
        'OnTurnError Trace',
        `${ error }`,
        'https://www.botframework.com/schemas/error',
        'TurnError'
    );

    // Send a message to the user
    await context.sendActivity('The bot encountered an error or bug.');
    await context.sendActivity('To continue to run this bot, please fix the bot source code.');
};

// Create the main dialog.
const myBot = new TeamsBot();

// Listen for incoming requests.
server.post('/api/messages', (req, res) => {
    adapter.processActivity(req, res, async (context) => {
        // Route to main dialog.
        await myBot.run(context);
    });
});

server.get('/api/notify', async (req, res) => {
    MicrosoftAppCredentials.trustServiceUrl('https://smba.trafficmanager.net/uk/');
    fs.readFile('data.json', async (err, data) => {
        var conversations = JSON.parse(data).conversations;
        var credentials = new MicrosoftAppCredentials(process.env.MicrosoftAppId, process.env.MicrosoftAppPassword);
        var client = new ConnectorClient(credentials, {baseUri: 'https://smba.trafficmanager.net/uk/'});
        conversations.forEach(async conversation => {
            
            var activityResponse = await client.conversations.sendToConversation(conversation.activity.conversation.id, {
                type: 'message',
                from: {id: process.env.MicrosoftAppId},
                text: 'heres some text.'
            });

            // var ref = TurnContext.getConversationReference(conversation.activity);
            // ref.user = conversation.user;
            // await adapter.continueConversation(ref, async turnContext => {
            //     MicrosoftAppCredentials.trustServiceUrl('https://smba.trafficmanager.net/uk/');
            //     await turnContext.sendActivity('There was an error');
            // });
        }); 
        res.send(conversations);
    });
});


