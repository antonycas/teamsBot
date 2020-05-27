// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { TeamsActivityHandler, TeamsInfo } = require('botbuilder');
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
                fs.readFile('conversations.json', (err, data) => {
                    if(err) {
                        console.log(err)
                    } else {
                        var obj = JSON.parse(data);
                        obj.conversations.push(details);
                        var json = JSON.stringify(obj);
                        fs.writeFile('conversations.json', json, 'utf8', () => {})
                    };
                });
            } 
            // By calling next() you ensure that the next BotHandler is run.
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

