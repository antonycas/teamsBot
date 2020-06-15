// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { TeamsActivityHandler, TeamsInfo, CardFactory, MessageFactory } = require('botbuilder');
const { ConnectorClient, MicrosoftAppCredentials } = require('botframework-connector');
const { Template } = require('adaptivecards-templating');
const fs = require('fs');

class TeamsBot extends TeamsActivityHandler {
    constructor() {
        super();

        this.onConversationUpdate(async (context, next) => {
            if(context.activity.membersAdded.length > 1) {
                let { membersAdded } = context.activity;
                fs.readFile(process.env.dataFile, (err, data) => {
                    let JSONData = JSON.parse(data);
                    if(err) {
                        console.log(err)
                    } else {
                        membersAdded.forEach(addedMember => {
                            if(!JSONData.users.some(user => user.id === addedMember.id )) {
                                JSONData.users.push(addedMember);     
                            }    
                        })
                        let toWrite = JSON.stringify(JSONData);
                        fs.writeFile(process.env.dataFile, toWrite, 'utf8', () => {})
                    }
                });
                await next();
            }
        })

        this.onEvent(async (context, next) => {
            if(context.activity.name === 'error') {
                var templatePayload = {
                    "type": "AdaptiveCard",
                    "version": "1.0",
                    "body": [
                        {
                            "type": "TextBlock",
                            "text": "${incidentId}",
                            "color": "attention",
                            "size": "large",
                            "weight": "bolder",
                            "spacing": "none"
                        },
                        {
                            "type": "TextBlock",
                            "text": "${date}",
                            "isSubtle": true,
                            "spacing": "none"
                        },
                        {
                            "type": "TextBlock",
                            "text": "${hostName}"
                        },
                        {
                            "type": "TextBlock",
                            "text": "${hostAddress}"
                        },
                        {
                            "type": "TextBlock",
                            "text": "${hostData}"
                        },
                        {
                            "type": "TextBlock",
                            "text": "${serviceName}"
                        },
                        {
                            "type": "TextBlock",
                            "text": "${serviceData}"
                        },
                        {
                            "type": "TextBlock",
                            "text": "${serviceStatus}"
                        }
                    ]
                }
                
                var template = new Template(templatePayload);
                let { data } = context.activity;
                var expandedCard = template.expand({
                    $root: {
                        incidentId: `ERROR: ${data.incidentId}`,
                        hostName: `Host Name: ${data.hostName}`,
                        hostAddress: `Host Address: ${data.hostAddress}`,
                        hostData: `Host Data: ${data.hostData}`,
                        serviceName: `Service Name: ${data.serviceName}`,
                        serviceData: `Service Data: ${data.serviceData}`,
                        serviceStatus: `Service Status: ${data.serviceStatus}`,
                        date: data.date
                    }
                });
                
                const card = CardFactory.adaptiveCard(expandedCard);
                var activity =  MessageFactory.attachment(card);
                
                MicrosoftAppCredentials.trustServiceUrl('https://smba.trafficmanager.net/uk/');
                var credentials = new MicrosoftAppCredentials(process.env.MicrosoftAppId, process.env.MicrosoftAppPassword);
                var client = new ConnectorClient(credentials, {baseUri: 'https://smba.trafficmanager.net/uk/'});

                var conversationParams = {
                    channelData: {
                        teamsChannelId: context.activity.teamsChannelId
                    },
                    activity: activity
                }

                const initialConversation = await client.conversations.createConversation(conversationParams);

                let usersToNotify = context.activity.usersToNotify;
                fs.readFile(process.env.dataFile, (err, data) => {
                    let JSONData = JSON.parse(data);
                    if(err) {
                        console.log(err)
                    } else {
                        usersToNotify.forEach(async user => {
                            let mention = {
                                mentioned: JSONData.users.filter(u =>  u.aadObjectId === user.id )[0],
                                text: `<at>${ user.displayName }</at>`,
                                type: 'mention'
                            };
                            
                            activity = MessageFactory.text(`<at>${ user.displayName }</at>`);
                            activity.entities = [mention];
                            console.log(activity, initialConversation, mention, mention.mentioned)
                            await client.conversations.sendToConversation(initialConversation.id, activity);
                        });  
                    }
                });
            }
            await next();
        }); 
        
    }
}

module.exports.TeamsBot = TeamsBot; 

