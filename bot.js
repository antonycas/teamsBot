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
                
                // write details to file
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
            
            await context.sendActivity('Hello. Send a message to the bot to begin listening for errors');
            // By calling next() you ensure that the next BotHandler is run.
            await next();
        });

        
        this.onMessage(async (context, next) => { 
            await next();
        });

        this.onEvent(async (context, next) => {
            this.readUserJSON(async data => {
                await data.conversations.forEach(async conversation => {
                    
                    
                    // const message = MessageFactory.text('There was an error');
                    // var ref = TurnContext.getConversationReference(conversation.activity);
                    // ref.user = conversation.user;
                    
                    // await context.adapter.createConversation(ref, async (t1) => {
                    //     const ref2 = TurnContext.getConversationReference(t1.activity);
                    //     await t1.adapter.continueConversation(ref2, async (t2) => {
                    //         await t2.sendActivity(message)
                    //     })
                    // });
                })
            });
            
            await context.sendActivity(''); 
            await next();
        });
    }

    async readUserJSON(callback) {
        await fs.readFile('data.json', async (err, data) => {
            if(err) {
                console.log(err)
            } else  {context
                await callback(JSON.parse(data));
            }
        });
    }
}

module.exports.TeamsBot = TeamsBot; 

