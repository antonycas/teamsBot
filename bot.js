// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { TeamsActivityHandler, MessageFactory, TeamsInfo, TurnContext } = require('botbuilder');
const { ConnectorClient, MicrosoftAppCredentials } = require('botframework-connector');
const { TeamsContext } = require('botbuilder-teams');
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
    
                fs.readFile('data.json', (err, data) => {
                    if(err) {
                        console.log(err)
                    } else {
                        var obj = JSON.parse(data);
                        obj.conversations.push(details);
                        var json = JSON.stringify(obj);
                        fs.writeFile('data.json', json, 'utf8', () => {})
                    };
                });
            } 
            
            
            // By calling next() you ensure that the next BotHandler is run.
            await next();
        });
        
        this.onMessage(async (context, next) => {
            
            this.readUserJSON(async data => {
                console.log(data.conversations);
                await data.conversations.forEach(async conversation => {
                    const message = MessageFactory.text('There was an error');
                    var ref = TurnContext.getConversationReference(conversation.activity);
                    ref.user = conversation.user;
                    
                    console.log(context.activity.serviceUrl);
                    MicrosoftAppCredentials.trustServiceUrl(context.activity.serviceUrl);
                    await context.adapter.createConversation(ref, async (t1) => {
                        const ref2 = await TurnContext.getConversationReference(t1.activity);
                        await t1.adapter.continueConversation(ref2, async (t2) => {
                            await t2.sendActivity(message)
                        })
                    });
                })
                // await this.messageUsersInConversations(data.conversations, context);
            });
            
            await context.sendActivity('its working'); 
            await next();
        });
    }

    async messageUsersInConversations(conversations, context) {
        await conversations.forEach(async conversation => {
            const message = MessageFactory.text('There was an error');
            var ref = await TurnContext.getConversationReference(conversation.activity);
            ref.user = conversation.user;
                    
            await context.adapter.createConversation(ref, async (t1) => {
                const ref2 = TurnContext.getConversationReference(t1.activity);
                await t1.adapter.continueConversation(ref2, async (t2) => {
                    await t2.sendActivity(message)
                })
            });
        })
    }

    async readUserJSON(callback) {
        await fs.readFile('data.json', async (err, data) => {
            if(err) {
                console.log(err)
            } else  {
                await callback(JSON.parse(data));
            }
        });
    }

    // async sendProactiveMessage(user, context) {
    //     context.activity.serviceUrl = 'https://smba.trafficmanager.net/uk/';
    //     MicrosoftAppCredentials.trustServiceUrl(context.activity.serviceUrl);
    //     const credentials = new MicrosoftAppCredentials(process.env.MicrosoftAppId, process.env.MicrosoftAppPassword);
    //     const connector = new ConnectorClient(credentials,{ baseUri: context.activity.serviceUrl });
    //     const teamsCtx = TeamsContext.from(context);
    //     const parameters = {
    //         members: [
    //             {id: user.id}
    //         ],
    //         channelData: {
    //             tenant: {
    //                 id: teamsCtx.tenant.id
    //             }
    //         }
    //     }
        
    //     const conversationResource = await connector.conversations.createConversation(parameters);
    //     const message = MessageFactory.text('This is a proactive message');
    //     await connector.conversations.sendToConversation(conversationResource.id, message);        
    // }
}

module.exports.TeamsBot = TeamsBot; 

