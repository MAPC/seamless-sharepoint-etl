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
COLUMN_VALUES = ['vendor', 'description', 'picker_erk', 'charge code', 'receipt_qFP']

logger = Logger.new("#{APP_ROOT}/log/seamless-to-sharepoint.log")
logger.level = Logger::INFO

# Create the string to sign according to the following pseudo-grammar
#
# StringToSign = HTTPVerb + "+" +
#                HTTPRequestURI + "+" +
#                <timestamp>
# See: http://developers.seamlessdocs.com/v1.2/docs/signing-requests#signature-base
#
def seamless_api_signature(request_uri, request_method, timestamp)
  key = SETTINGS['seamless']['secret']
  data = request_method + '+' +
         URI(request_uri).path.gsub(%r{/api}, '') + '+' +
         timestamp
  OpenSSL::HMAC.hexdigest('SHA256', key, data)
end

def microsoft_oauth2_token
  # See: https://docs.microsoft.com/en-us/graph/auth-v2-user#2-get-authorization
  # https://docs.microsoft.com/en-us/azure/active-directory/develop/quickstart-v2-python-daemon
  client = OAuth2::Client.new(SETTINGS['microsoft']['client_id'],
                              SETTINGS['microsoft']['client_secret'],
                              authorize_url: 'https://login.microsoftonline.com/c75d8168-fa8e-4753-8aef-55111ae727bd/oauth2/v2.0/authorize',
                              token_url: 'https://login.microsoftonline.com/c75d8168-fa8e-4753-8aef-55111ae727bd/oauth2/v2.0/token')

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

  JSON.parse(response.body)
      .to_hash['value']
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

  # Get column machine names from seamless
  columns = JSON.parse(response.body).to_hash['columns']
                                     .values
                                     .filter { |hash| COLUMN_VALUES.include?(hash['printable_name']) }

  sorted_columns = COLUMN_VALUES.map { |value| columns.select { |column| column['printable_name'] == value } }
                                .map { |hash| hash[0]['column_id'] }

  values_formatted_for_microsoft = []
  JSON.parse(response.body)['items'].each do |item|
    item_data = []
    sorted_columns.each do |column|
      item_data << item['application_data'][column]
    end
    values_formatted_for_microsoft << item_data
  end
  return values_formatted_for_microsoft
end

def add_seamless_data_to_sharepoint(data)
  connection = Faraday.new 'https://graph.microsoft.com' do |conn|
    conn.headers['Authorization'] = "Bearer #{microsoft_oauth2_token.token}"
    conn.adapter Faraday.default_adapter
  end
  # POST /workbook/worksheets/{id|name}/tables/{id|name}/rows/add
  response = connection.post('/v1.0/sites/mapc365.sharepoint.com,86503781-e6fa-4516-abbf-879f74eaac01,4a7dbaee-d756-4a4b-b1f5-897b4f3c31a2/drive/root:/Digital%20Municipal%20Work/Seamless%20API%20Test.xlsx:/workbook/worksheets/Sheet1/tables/Table1/rows/add') do |req|
    req.body = {
                "index": nil,
                "values": data
              }.to_json
  end
end

def receipt_qfp_to_i(receipt_qFP)
  receipt_qFP.delete_prefix('U')
             .chop
             .to_i
end



puts receipt_qfp_to_i(get_last_entry_from_sharepoint)
# add_seamless_data_to_sharepoint(get_seamless_form_data)
# TODO: push the new rows to SharePoint

