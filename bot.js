// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { TeamsActivityHandler, CardFactory, MessageFactory } = require('botbuilder');
const { ConnectorClient, MicrosoftAppCredentials } = require('botframework-connector');
const { Template } = require('adaptivecards-templating');
const { BlobStorage } = require('botbuilder-azure')


class TeamsBot extends TeamsActivityHandler {
    constructor() {
        super();

        this.serviceUrl = 'https://smba.trafficmanager.net/uk/';
        this.credentials = new MicrosoftAppCredentials(process.env.MicrosoftAppId, process.env.MicrosoftAppPassword);
        this.client = new ConnectorClient(this.credentials, {baseUri: this.serviceUrl });
        this.storage = new BlobStorage({
            containerName: process.env.BLOB_NAME,
            storageAccountOrConnectionString: process.env.BLOB_STRING
        });

        this.onMessage(async (context, next) => {
            const mention = {
                mentioned: context.activity.from,
                text: `<at>${context.activity.from}</at>`,
                type: 'mention'
            }
            const activity = MessageFactory.text(`Hello ${mention.text}`);
            await context.sendActivity(activity)
            await next();
        })

        this.onConversationUpdate(async (context, next) => {
            if(context.activity.membersAdded.length > 1) {
                let { membersAdded } = context.activity;
                let storeItems = await this.storage.read(["users"]);
                let users = storeItems["users"]
                if(typeof (users) != 'undefined') {
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
                
                let storeItems = await this.storage.read(["users", "incidents"])
                let { users, incidents } = storeItems;
                
                if (typeof (incidents) === 'undefined') { storeItems.incidents = [] }
                if (incidents.some(i => { return i.id === context.activity.data.incidentId })) {
                    console.log('Error already exists')
                } else {
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
                    var conversationParams = {
                        channelData: {
                            teamsChannelId: context.activity.teamsChannelId
                        },
                        activity: activity
                    }
                    
                    MicrosoftAppCredentials.trustServiceUrl(this.serviceUrl);
                    const initialConversation = await this.client.conversations.createConversation(conversationParams);
                    
                    storeItems.incidents.push({
                        id: context.activity.data.incidentId,
                        conversation: initialConversation 
                    })
                    this.saveData(storeItems)

                    let usersToNotify = users.filter(u => context.activity.usersToNotify.some(user => u.aadObjectId == user.id))
                    this.getUserDisplayNamesFromContext(context, usersToNotify);
                    await this.notifyUsersOfActivity(usersToNotify, initialConversation.id, initialConversation.activityId)
                }
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

                let storeItems = await this.storage.read(["users","incidents"]);
                let { users, incidents } = storeItems;
                let initialIncident = incidents.filter(i => i.id === context.activity.data.incidentId)[0]
                let { conversation } = initialIncident
                MicrosoftAppCredentials.trustServiceUrl(this.serviceUrl);
                await this.client.conversations.replyToActivity(conversation.id, conversation.activityId, activity)

                let usersToNotify = users.filter(u => context.activity.usersToNotify.some(user => u.aadObjectId == user.id))
                this.getUserDisplayNamesFromContext(context, usersToNotify);
                await this.notifyUsersOfActivity(usersToNotify, conversation.id, conversation.activityId) 
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

    async notifyUsersOfActivity(usersToNotify, conversationId, activityId) {
        usersToNotify.forEach(async user => {
            let mention = {
                mentioned: user,
                text: `<at>${user.displayName}</at>`,
                type: 'mention'
            }
            let activity = MessageFactory.text(`${user.displayName} ${mention.text}`)
            activity.entities = [mention]
            await this.client.conversations.replyToActivity(conversationId, activityId, activity)
        })
    }

    getUserDisplayNamesFromContext(context, users) {
        users.forEach(user => {
            user.displayName = context.activity.usersToNotify.filter(u=> { return u.id === user.aadObjectId })[0].displayName
        })
    }
}

module.exports.TeamsBot = TeamsBot; 

