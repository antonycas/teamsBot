require('uri')

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

	def get_channels(group_id)
		command = "curl -H 'Authorization:#{access_token}' 'https://graph.microsoft.com/v1.0/teams/#{group_id}/channels'"
		JSON.parse(`#{command}`)['value']
	end

	def create_team(members=[], display_name)
		request_body = {
			displayName: display_name,
			mailEnabled: true,
			mailNickname: 'naemon errors',
			securityEnabled: true,
			members: members
		}.to_json
		command = "curl -H 'Authorization:#{access_token}' -H 'Content-Type: application/json' -d #{request_body} 'https://graph.microsoft.com/v1.0/groups'"
		JSON.parse(`#{command}`)
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
		puts "getting app list"
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