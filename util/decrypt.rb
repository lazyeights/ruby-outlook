# decrypt.rb
# Copyright (c) 2009 David B. Conrad

require '../lib/encryption'
include Pst

if ARGV.size < 1 then Process.exit end

File.open(ARGV.first, "rb") do |file|

  @data = file.read
  @cleartext = CompressibleEncryption.decrypt @data
  
end

File.open("#{ARGV.first}.clear", "wb") do |file|
  
  file.write(@cleartext)
  
end