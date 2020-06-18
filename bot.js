// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { TeamsActivityHandler, CardFactory, MessageFactory } = require('botbuilder');
const { ConnectorClient, MicrosoftAppCredentials } = require('botframework-connector');
const { Template } = require('adaptivecards-templating');
const fs = require('fs');
const { BlobStorage } = require('botbuilder-azure')


class TeamsBot extends TeamsActivityHandler {
    constructor() {
        super();
        this.storage = new BlobStorage({
            containerName: process.env.BLOB_NAME,
            storageAccountOrConnectionString: process.env.BLOB_STRING
        });
        
        this.onConversationUpdate(async (context, next) => {
            if(context.activity.membersAdded.length > 1) {
                let { membersAdded } = context.activity;
                let storeItems = await this.storage.read(["users"]);
                let users = storeItems["users"]
                if(typeof (users) != 'undefined') {
                    console.log('got here')
                    membersAdded.forEach(adddedMember => {
                        if(!storeItems.users.some(user => user.id === adddedMember.id)) {
                            storeItems.users.push(adddedMember)
                        }
                    });
                    this.saveData(storeItems);
                } else {
                    storeItems["users"] = membersAdded
                    this.saveData(storeItems);
                } 
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

                let storeItems = await this.storage.read(["users", "incidents"])
                let { users, incidents } = storeItems;
                let usersToNotify = users.filter(u => context.activity.usersToNotify.some(user => u.aadObjectId == user.id))

                if (typeof (incidents) === 'undefined') { storeItems.incidents = [] }
                storeItems.incidents.push({
                    id: context.activity.data.incidentId,
                    conversation: initialConversation 
                }) 
                this.saveData(storeItems)

                usersToNotify.forEach(async user => {
                    let mention = {
                        mentioned: user,
                        text: `<at>${user.displayName}</at>`,
                        type: 'mention'
                    }
                    activity = MessageFactory.text(`<at>${user.displayName}</at>`);
                    activity.entities = [mention]
                    await client.conversations.replyToActivity(initialConversation.id, initialConversation.activityId, activity)
                })
            } else if(context.activity.name === 'resolved') {
                var templatePayload = {
                    "type": "AdaptiveCard",
                    "version": "1.0",
                    "body": [
                        {
                            "type": "TextBlock",
                            "text": "${incidentId}",
                            "color": "good",
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
                        incidentId: `${data.incidentId}(RESOLVED)`,
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

                let storeItems = await this.storage.read(["incidents"]);
                let { incidents } = storeItems;
                let initialIncident = incidents.filter(i => i.id === context.activity.data.incidentId)[0]
                let { conversation } = initialIncident
                await client.conversations.replyToActivity(conversation.id, conversation.activityId, activity) 
            }
            await next();
        }); 
        
    }

    async saveData(data) {
        try {
            await this.storage.write(data)
        } catch(err) {
            console.log(err)
        }
    }

    saveToDataFile(data) {
        fs.writeFile(process.env.dataFile, JSON.stringify(data), err => {
            if(err) { console.log(err) } 
        })
    }
}

module.exports.TeamsBot = TeamsBot; 

