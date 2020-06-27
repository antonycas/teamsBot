require('uri')
require('cgi')

class MSGraphClient

	def initialize(client_id, client_secret, aad_name)
		@client_secret = client_secret
		@client_id = client_id
		@aad_name = aad_name
		@authorization = get_authorization
	end

	def get_authorization(grant_type="client_credentials", resource="https://graph.microsoft.com")
		command = "curl -H 'Content-Type: application/x-www-form-urlencoded' -X POST 'https://login.microsoftonline.com/#{@aad_name}/oauth2/token' -d client_id=#{@client_id} -d client_secret=#{@client_secret} -d grant_type=#{grant_type} -d resource=#{resource}"
		JSON.parse(`#{command}`)
	end

	def get_groups
		command = "curl -H 'Authorization:#{access_token}' 'https://graph.microsoft.com/v1.0/groups'"
		JSON.parse(`#{command}`)['value']
	end

	def get_members(group_id)
		command = "curl -H 'Authorization:#{access_token}' 'https://graph.microsoft.com/v1.0/groups/#{group_id}/members'"
		JSON.parse(`#{command}`)['value']
	end

	def get_user(user_pricipal_name)
		JSON.parse(`curl -H 'Authorization: #{access_token}' 'https://graph.microsoft.com/v1.0/users/#{user_pricipal_name}'`)
	end

	def add_member_to_team(team_id, member_id)
		url = "https://graph.microsoft.com/v1.0/groups/#{team_id}/members/%24ref"
		request_body = { '@odata.id': "https://graph.microsoft.com/v1.0/directoryObjects/#{member_id}" }.to_json
		`curl -H 'Authorization: #{access_token}' -H 'Content-Type: application/json' -d '#{request_body}' #{url}`
	end

	def list_channels(group_id)
		command = "curl -H 'Authorization:#{access_token}' 'https://graph.microsoft.com/v1.0/teams/#{group_id}/channels'"
		JSON.parse(`#{command}`)['value']
	end

	def update_channel(group_id, channel_id, channel_data)
		`curl -H 'Authorization: #{access_token}' -H 'Content-Type: application/json' -X PATCH 'https://graph.microsoft.com/v1.0/teams/#{group_id}/channels/#{channel_id}' -d '#{channel_data}'`
	end

	def delete_channel(team_id, channel_id)
		`curl -H 'Authorization: #{access_token}' -X DELETE 'https://graph.microsoft.com/v1.0/teams/#{team_id}/channels/#{channel_id}'`
	end

	def create_group(owner_id, display_name, members=[])
		request_body = {
			"displayName": display_name,
			"mailEnabled": false,
			"mailNickname": 'naemon_errors',
			"securityEnabled": true,
			"owners@odata.bind": ["https://graph.microsoft.com/v1.0/users/#{owner_id}"],
			"members@odata.bind": ["https://graph.microsoft.com/v1.0/users/#{owner_id}"]
		}
		command = "curl -H 'Authorization:#{access_token}' -H 'Content-Type: application/json' -d #{request_body.to_json.to_json} 'https://graph.microsoft.com/v1.0/groups'"
		JSON.parse(`#{command}`)
	end

	def create_team(group_id)
		request_body = {  
			memberSettings: {
				allowCreateUpdateChannels: true
			},
			messagingSettings: {
				allowUserEditMessages: true,
				allowUserDeleteMessages: true
			},
			funSettings: {
				allowGiphy: true,
				giphyContentRating: "strict"
			}
		}.to_json.to_json 
		command = "curl -H 'Authorization: #{access_token}' -H 'Content-Type: application/json' -d #{request_body} -X PUT 'https://graph.microsoft.com/v1.0/groups/#{group_id}/team'"
		JSON.parse `#{command}`
	end

	def create_teams_channel(team_id, display_name, description)
		request_body = {
			displayName: display_name,
			description: description
		}.to_json.to_json # needs calling twice for some reason
		command = "curl -H 'Authorization: #{access_token}' -H 'Content-Type: application/json' -d #{request_body} 'https://graph.microsoft.com/v1.0/teams/#{team_id}/channels'"
		JSON.parse(`#{command}`)
	end

	def get_installed_apps(user_id)
		href = "https://graph.microsoft.com/beta/users/#{user_id}/teamwork/installedApps?$expand=teamsAppDefinition"
		command = "curl -H 'Authorization: #{access_token}' '#{href}'"
		JSON.parse(`#{command}`)['value']
	end

	def install_teams_app(user_id, app_id)
		puts "installing app for user with id #{user_id}"
		request_body = {
			"teamsApp@odata.bind": "https://graph.microsoft.com/beta/appCatalogs/teamsApps/#{app_id}"
		}.to_json.to_json
		command = "curl -H 'Authorization: #{access_token}' -H 'Content-Type: application/json' -d #{request_body} 'https://graph.microsoft.com/beta/users/#{user_id}/teamwork/installedApps'"
		`#{command}`
	end

	private

	def access_token
		@authorization["access_token"]
	end
end