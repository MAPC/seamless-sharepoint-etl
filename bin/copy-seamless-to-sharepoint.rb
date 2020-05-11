#!/usr/bin/env ruby
require 'rubygems'
require 'bundler/setup'
require 'faraday'
require 'yaml'
require 'dotenv'
require 'openssl'
require 'logger'

logger = Logger.new('../log/seamless-to-sharepoint.log')
logger.level = Logger::INFO

settings = YAML.load_file('../config/settings.yml')

# TODO: Get current state/most recent item(s) from the Microsoft Sharepoint document

# TODO: Read input data from the seamless API based on above

# TODO: Push input data to Microsoft Sharepoint Document table

# TODO: Translate signature formula code from Paw JS file to Ruby method
def seamless_api_signature(date)
  key = "key"
  data = "message-to-be-authenticated"
  mac = OpenSSL::HMAC.hexdigest("SHA256", key, data)
end
