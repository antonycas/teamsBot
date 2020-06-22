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

def todays_channel(ms_graph_client, team_id)
  date_string = Time.new.strftime("%d_%m_%y") 
  channels = ms_graph_client.list_channels(team_id)
  channel = channels.find { |c| c['displayName'].split(' ').first == date_string }
  (channel.nil?) ? ms_graph_client.create_teams_channel(team_id, date_string, date_string) : channel 
end

def start_conversation_with_bot(direct_channel_secret)
  command = "curl -X POST -H 'Authorization: Bearer #{direct_channel_secret}' -H 'Content-Type: application/json' -d '' 'https://directline.botframework.com/v3/directline/conversations'"
  JSON.parse(`#{command}`)
end

def send_activity_to_bot(direct_channel_secret, activity, conversation_id)
  command = "curl -H 'Authorization: Bearer #{direct_channel_secret}' -H 'Content-Type: application/json' -d '#{activity}' 'https://directline.botframework.com/v3/directline/conversations/#{conversation_id}/activities'"
  `#{command}`
end

client_id = "269ab60d-9dd1-4a29-a4b5-807f788ada90" # id of registered app on azure portal
client_secret = "Dt~u~BtwGWJh8u-~gM9u0-8XN-o04-B8S1" # secret key generated in dashboard of registered app
aad_name = "antcasdev.onmicrosoft.com" # microsoft aad_name, i.e antcasdev.onmicrosoft.com
direct_channel_secret = ARGV[0] 
data = JSON.parse(ARGV[1])
ms_graph_client = MSGraphClient.new(client_id, client_secret, aad_name)

users_to_notify = []
user_emails = data['usersToNotify']
user_emails.each {|e| users_to_notify << ms_graph_client.get_user(e) }

if data['status'] == nil
  abort('No status provided.')
elsif data['status'] == 'error'
  teams_app_id = '500f16aa-318c-4bdc-a8ae-05855567d31a'
  users_to_notify.each do |u|
    installed_apps = ms_graph_client.get_installed_apps(u['id'])
    new_app_installation = ms_graph_client.install_teams_app(u['id'], teams_app_id) if installed_apps.none? {|a| a['teamsAppDefinition']['teamsAppId'] == teams_app_id }
  end
  
  team_id = '9b41ef88-1ea6-4ffa-915d-81e27c0e3ae1'
  channel = todays_channel(ms_graph_client,team_id)

  conversation = start_conversation_with_bot(direct_channel_secret)
  activity = {
    type: 'event',
    name: 'error',
    from: {
      id: 'errorBot'
    },
    data: data,
    teamsChannelId: channel['id'],
    usersToNotify: users_to_notify
  }.to_json 
  send_activity_to_bot(direct_channel_secret, activity, conversation['conversationId'])
elsif data['status'] == 'resolved'
  conversation = start_conversation_with_bot(direct_channel_secret)
  activity = {
    type: 'event',
    name: 'resolved',
    from: {
      id: 'errorBot'
    },
    data: data,
    usersToNotify: users_to_notify
  }.to_json
  send_activity_to_bot(direct_channel_secret, activity, conversation['conversationId'])
end
