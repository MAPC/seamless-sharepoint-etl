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

APP_ROOT = File.expand_path("..", __dir__)
Dotenv.load("#{APP_ROOT}/.env")
template = ERB.new File.new("#{APP_ROOT}/config/settings.yml").read
SETTINGS = YAML.load template.result(binding)

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

# TODO: Get current state/most recent item(s) from the Microsoft Sharepoint document

def get_seamless_form_data(form_id = 'CO20041000144715117')
  request_uri = "https://mapc.seamlessdocs.com/api/form/#{form_id}/pipeline"
  timestamp = Time.now.to_i.to_s
  signature = seamless_api_signature(request_uri, 'GET', timestamp)

  response = Faraday.get(request_uri) do |request|
    request.headers['AuthDate'] = timestamp
    request.headers['Authorization'] = "HMAC-SHA256 api_key=#{SETTINGS['seamless']['api_key']} signature=#{signature}"
  end
end

# TODO: Push input data to Microsoft Sharepoint Document table
