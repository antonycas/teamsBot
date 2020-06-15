#!/usr/bin/ruby
require 'json'
require './ms_graph_client'

def display_name(string)
  s = string
  illegal_chars = "#%&*{}/\:<>?+|'"
  illegal_chars.each_char {|c|
    s.gsub!(c,'_')
  }
  s
end

client_id = "269ab60d-9dd1-4a29-a4b5-807f788ada90" # id of registered app on azure portal
client_secret = "Dt~u~BtwGWJh8u-~gM9u0-8XN-o04-B8S1" # secret key generated in dashboard of registered app
aad_name = "antcasdev.onmicrosoft.com" # microsoft aad_name, i.e antcasdev.onmicrosoft.com
direct_channel_secret = ARGV[0] 
data = JSON.parse(ARGV[1])
ms_graph_client = MSGraphClient.new(client_id, client_secret, aad_name)

# list all groups 
groups = ms_graph_client.get_groups
# select a group based on display name
group = groups.find { |g| g['displayName'] == 'antcasdev'}
# return the selected members
members = ms_graph_client.get_members(group['id'])

user_emails = ['antcas@antcasdev.onmicrosoft.com', 'testuser@antcasdev.onmicrosoft.com']
users_to_notify = members.select {|m| user_emails.include?(m['mail']) }

teams_app_id = '500f16aa-318c-4bdc-a8ae-05855567d31a'
users_to_notify.each {|u|
  installed_apps = ms_graph_client.get_installed_apps(u['id'])
  new_app_installation = ms_graph_client.install_teams_app(u['id'], teams_app_id) if installed_apps.none? {|a| a['teamsAppDefinition']['teamsAppId'] == teams_app_id }
}

team_id = "dc554e46-e99f-4f17-8261-fe6eb07c721d"
display_string = "#{data['date']} ID #{data['incidentId']}"
dn = display_name(display_string)
description = dn
new_channel = ms_graph_client.create_teams_channel(team_id, dn, description)
pp new_channel
# # start a conversation with the bot
command = "curl -X POST -H 'Authorization: Bearer #{direct_channel_secret}' -H 'Content-Type: application/json' -d '' 'https://directline.botframework.com/v3/directline/conversations'"
conversation = JSON.parse(`#{command}`)

error_message = "There has been an error" 
# send activity to started conversation
conversation_id = conversation['conversationId']
activity = {
  text: 'hadsjk',
  textFormat: 'plain',
  type: 'event',
  name: 'error',
  from: {
    id: 'errorBot'
  },
  data: data,
  teamsChannelId: new_channel['id'],
  usersToNotify: users_to_notify,
  userEmails: ['antcas@antcasdev.onmicrosoft.com', 'romancollyer@antcasdev.onmicrosoft.com', 'testuser@antcasdev.onmicrosoft.com']
}.to_json


command = "curl -H 'Authorization: Bearer #{direct_channel_secret}' -H 'Content-Type: application/json' -d '#{activity}' 'https://directline.botframework.com/v3/directline/conversations/#{conversation_id}/activities'"
x = `#{command}`
pp x