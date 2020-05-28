// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { TeamsActivityHandler, TeamsInfo } = require('botbuilder');
const { ConnectorClient, MicrosoftAppCredentials } = require('botframework-connector');
const fs = require('fs');

class TeamsBot extends TeamsActivityHandler {
    constructor() {
        super();

        this.onMembersAdded(async (context, next) => {
            
            // only save the members details if the member is added through teams
            if(context.activity.channelId === 'msteams') {
                var members = await TeamsInfo.getMembers(context);
                // select any members who are not bots
                var users = members.filter(member => { return member.name.toLowerCase() !== "bot" });
                var details = {
                    user: users[0],
                    activity: context.activity
                };
                
                // write details to file
                fs.readFile(process.env.dataFile, (err, data) => {
                    if(err) {
                        console.log(err)
                    } else {
                        var obj = JSON.parse(data);
                        obj.conversations.push(details);
                        var json = JSON.stringify(obj);
                        fs.writeFile(process.env.dataFile, json, 'utf8', () => {})
                    };
                });
            } 
            // By calling next() you ensure that the next BotHandler is run.
            await next();
        });

        this.onMessage(async (context, next) => {
            // crete activity to be sent to conversation
            var activity = {
                type: 'message',
                from: {id: process.env.MicrosoftAppId},
                text: `${context.activity.from.name}: ${context.activity.text}`
            }

            var fromId = context.activity.from.id

            fs.readFile(process.env.dataFile, async (err, data) => {
                if(err) {
                    return console.log(err)
                } else {
                    var conversations = JSON.parse(data).conversations;
                    var filteredConversations = conversations.filter(conversation => { return conversation.user.id !== fromId });
                    this.sendToConversations(filteredConversations, activity);
                }
            })

            await next();
        });


        this.onEvent(async (context, next) => {
            if(context.activity.name === 'error') {
                var activity = {
                    type: 'message',
                    from: {id: process.env.MicrosoftAppId},
                    text: context.activity.text
                }

                fs.readFile(process.env.dataFile, async (err, data) => {
                    if(err) {
                        return console.log(err)
                    } else {
                        var conversations = JSON.parse(data).conversations;
                        this.sendToConversations(conversations, activity);
                    }
                })
            }
            await next();
        }); 

    }

    sendToAllConversations(activity) {
        MicrosoftAppCredentials.trustServiceUrl('https://smba.trafficmanager.net/uk/');
        fs.readFile(process.env.dataFile, async (err, data) => {
            var conversations = JSON.parse(data).conversations;
            var credentials = new MicrosoftAppCredentials(process.env.MicrosoftAppId, process.env.MicrosoftAppPassword);
            var client = new ConnectorClient(credentials, {baseUri: 'https://smba.trafficmanager.net/uk/'});
            conversations.forEach(async conversation => {      
                await client.conversations.sendToConversation(conversation.activity.conversation.id, activity);
            }); 
        });
    }

    sendToConversations(conversations, activity) {
        MicrosoftAppCredentials.trustServiceUrl('https://smba.trafficmanager.net/uk/');
        var credentials = new MicrosoftAppCredentials(process.env.MicrosoftAppId, process.env.MicrosoftAppPassword);
        var client = new ConnectorClient(credentials, {baseUri: 'https://smba.trafficmanager.net/uk/'});
        conversations.forEach(async conversation => {      
            await client.conversations.sendToConversation(conversation.activity.conversation.id, activity);
        });
    }
}

module.exports.TeamsBot = TeamsBot; 

