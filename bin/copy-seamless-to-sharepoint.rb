#!/usr/bin/env ruby
require 'dotenv'
require 'rubygems'
require 'bundler/setup'
require 'faraday'
require 'yaml'
require 'openssl'
require 'logger'
require 'uri'
require 'erb'
require 'oauth2'
require 'pry-byebug'

APP_ROOT = File.expand_path("..", __dir__)
Dotenv.load("#{APP_ROOT}/.env")
template = ERB.new File.new("#{APP_ROOT}/config/settings.yml").read
SETTINGS = YAML.load template.result(binding)
# ENV['OAUTH_DEBUG'] = 'true'

logger = Logger.new("#{APP_ROOT}/log/seamless-to-sharepoint.log")
logger.level = Logger::INFO

# Create the string to sign according to the following pseudo-grammar
#
# StringToSign = HTTPVerb + "+" +
#                HTTPRequestURI + "+" +
#                <timestamp>
# See: http://developers.seamlessdocs.com/v1.2/docs/signing-requests#signature-base
#
# The HTTPRequestURI component is the HTTP absolute path component of the
# URI up to, but not including, the query string. If the HTTPRequestURI is
# empty, use a forward slash ( / ).
#
def seamless_api_signature(request_uri, request_method, timestamp)
  key = SETTINGS['seamless']['secret']
  data = request_method + '+' +
         URI(request_uri).path.gsub(%r{/api}, '') + '+' +
         timestamp
  OpenSSL::HMAC.hexdigest('SHA256', key, data)
end

def microsoft_oauth2_token
  client = OAuth2::Client.new(SETTINGS['microsoft']['client_id'],
                              SETTINGS['microsoft']['client_secret'],
                              authorize_url: 'https://login.microsoftonline.com/c75d8168-fa8e-4753-8aef-55111ae727bd/oauth2/v2.0/authorize',
                              token_url: 'https://login.microsoftonline.com/c75d8168-fa8e-4753-8aef-55111ae727bd/oauth2/v2.0/token')
  # See: https://docs.microsoft.com/en-us/graph/auth-v2-user#2-get-authorization
  # https://docs.microsoft.com/en-us/azure/active-directory/develop/quickstart-v2-python-daemon

  client.client_credentials.get_token(scope: 'https://graph.microsoft.com/.default')
end

# TODO: Figure out why ordering by Request-qFP column on server is not working.
def get_last_entry_from_sharepoint
  connection = Faraday.new 'https://graph.microsoft.com' do |conn|
    conn.headers['Authorization'] = "Bearer #{microsoft_oauth2_token.token}"
    conn.adapter Faraday.default_adapter
  end

  # GET /sites/{site-id}/drive/root:/{item-path}
  response = connection.get("/v1.0/sites/mapc365.sharepoint.com,86503781-e6fa-4516-abbf-879f74eaac01,4a7dbaee-d756-4a4b-b1f5-897b4f3c31a2/drive/root:/Digital%20Municipal%20Work/Seamless%20API%20Test.xlsx:/workbook/worksheets/Sheet1/tables/Table1/rows")

  puts JSON.parse(response.body)
            .to_hash['value']
            .sort_by { |item| item.values[4] }
            .last['values'][0][4]
end

def get_seamless_form_data(form_id = 'CO20041000144715117')
  request_uri = "https://mapc.seamlessdocs.com/api/form/#{form_id}/pipeline"
  timestamp = Time.now.to_i.to_s
  signature = seamless_api_signature(request_uri, 'GET', timestamp)

  response = Faraday.get(request_uri) do |request|
    request.headers['AuthDate'] = timestamp
    request.headers['Authorization'] = "HMAC-SHA256 api_key=#{SETTINGS['seamless']['api_key']} signature=#{signature}"
  end
end

def add_seamless_data_to_sharepoint
  # POST /workbook/worksheets/{id|name}/tables/{id|name}/rows/add
    # resp = connection.post('v1.0/me/drive/root:/api-test.xlsx:/workbook/worksheets/Sheet1/tables/add') do |req|
    #   req.params['limit'] = 100
    #   req.body = {
    #               "index": null,
    #               "values": [
    #                 [
    #                   "Luke Skywalker",
    #                   "luke@skywalker.com",
    #                   "test",
    #                   "value"
    #                 ]
    #               ]
    #             }.to_json
    # end
end

get_last_entry_from_sharepoint
# TODO: Push input data to Microsoft Sharepoint Document table
