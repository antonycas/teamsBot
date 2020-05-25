// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { TeamsActivityHandler, MessageFactory, TeamsInfo, TurnContext } = require('botbuilder');
const { ConnectorClient, MicrosoftAppCredentials } = require('botframework-connector');
const { TeamsContext } = require('botbuilder-teams');
// const fs = require('fs');

class TeamsBot extends TeamsActivityHandler {
    constructor() {
        super();

        this.onMembersAdded(async (context, next) => {
            var members = await TeamsInfo.getMembers(context);
            var details = {
                user: members[0],
                activity: context.activity
            };

            await context.sendActivity(JSON.stringify(details, null, 2));

            // fs.readFile('data.json', (err, data) => {
            //     if(err) {
            //         console.log(err)
            //     } else {
            //         obj = JSON.parse(data);
            //         obj.push(details);
            //         json = JSON.stringify(obj);
            //         fs.writeFile('data.json', json, 'utf8', () => {
            //             context.sendActivity('Hello. Activity and User details have been saved.');
            //         })
            //     };
            // });

            const membersAdded = context.activity.membersAdded;
            for (let cnt = 0; cnt < membersAdded.length; ++cnt) {
                if (membersAdded[cnt].name !== "Bot") {
                    await context.sendActivity(`${membersAdded[cnt].name} joined the chat`);
                }
            }
            // By calling next() you ensure that the next BotHandler is run.
            await next();
        });
        
        this.onMessage(async (context, next) => {
            await context.sendActivity('its working');
            // context.activity.serviceUrl = 'https://smba.trafficmanager.net/uk/';
            // context.activity.users.forEach(user => this.sendProactiveMessage(user, context));
            // console.log('success');
            // fs.readFile('data.json', (err, data) => {
            //     console.log(data); 
            // })
            
            // var members = await TeamsInfo.getMembers(context);
            // context.sendActivity(JSON.stringify(members));        
            // // members.forEach(async teamMember => {
            //     const message = MessageFactory.text('Its working');
            //     var ref = TurnContext.getConversationReference({
            //         "membersAdded": [
            //         {
            //         "id": "28:269ab60d-9dd1-4a29-a4b5-807f788ada90"
            //         },
            //         {
            //         "id": "29:1ZkOXr1QW20VLzKY6NlnsZbMVnmBDtnp8oLEcZird_6g5Iry3LWB9UYomNrV32JPEKYCeE6KnsVdGIfjFimQsdA",
            //         "aadObjectId": "17498f3b-3830-49f3-9650-ce4d23ddecb5"
            //         }
            //         ],
            //         "type": "conversationUpdate",
            //         "timestamp": "2020-05-25T13:49:40.257Z",
            //         "id": "f:31415571-9f18-c240-fadb-70da291d28f6",
            //         "channelId": "msteams",
            //         "serviceUrl": "https://smba.trafficmanager.net/uk/",
            //         "from": {
            //         "id": "29:1ZkOXr1QW20VLzKY6NlnsZbMVnmBDtnp8oLEcZird_6g5Iry3LWB9UYomNrV32JPEKYCeE6KnsVdGIfjFimQsdA",
            //         "aadObjectId": "17498f3b-3830-49f3-9650-ce4d23ddecb5"
            //         },
            //         "conversation": {
            //         "conversationType": "personal",
            //         "tenantId": "d30f162e-47f1-411d-a1c1-7dd3526f0eef",
            //         "id": "a:1rknVIkQ6kiYklZbQA8asYy7Es4R1qruq8vZjXRSeRJ62VhxYWz8DRs4q_31QGJC7VR9HrWfNw-T9AqYbRmmCxio8epzflxhjuUrw1b2zwgWBnYnWysih3iQXirqYxOZl"
            //         },
            //         "recipient": {
            //         "id": "28:269ab60d-9dd1-4a29-a4b5-807f788ada90",
            //         "name": "antCasDevBot"
            //         },
            //         "channelData": {
            //         "tenant": {
            //         "id": "d30f162e-47f1-411d-a1c1-7dd3526f0eef"
            //         }
            //         }
            //         });
            //     ref.user = {"id":"29:1ZkOXr1QW20VLzKY6NlnsZbMVnmBDtnp8oLEcZird_6g5Iry3LWB9UYomNrV32JPEKYCeE6KnsVdGIfjFimQsdA","name":"antony castineiras","objectId":"17498f3b-3830-49f3-9650-ce4d23ddecb5","givenName":"antony","surname":"castineiras","email":"antcas@antcasdev.onmicrosoft.com","userPrincipalName":"antcas@antcasdev.onmicrosoft.com","tenantId":"d30f162e-47f1-411d-a1c1-7dd3526f0eef","userRole":"user","aadObjectId":"17498f3b-3830-49f3-9650-ce4d23ddecb5"};
                
            //     await context.adapter.createConversation(ref, async (t1) => {
            //         const ref2 = TurnContext.getConversationReference(t1.activity);
            //         await t1.adapter.continueConversation(ref2, async (t2) => {
            //             await t2.sendActivity(message)
            //         })
            //     });
            // });


            await next();
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

