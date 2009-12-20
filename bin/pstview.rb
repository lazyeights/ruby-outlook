# pstview - inspect Microsoft Outlook PST files
# Copyright (c) 2009 David B. Conrad

$LOAD_PATH << '../lib/pst'

require 'pst_file'
require 'mapi_types'
require 'mapi_message'

def to_hex value, type
  return "Nil" if value.nil?
  case type
  when 'PT_BINARY'  
    str = value.unpack("C#{value.length}").collect{ |c| "%02x" % c }.join #.gsub(/.{8}/) { |s| s+' | '}
  when 'PT_LONG'
    str = value.to_s(16)
  when 'PT_UNICODE'
    str = '"'+value+'"'
  when 'PT_BOOLEAN'
    str = value ? "True" : "False"
  when 'PT_SYSTIME'
    str = Time.at((value.unpack('Q').first - 116444736000000000) / 10000000).to_s
  else
    str = value
  end
  str = str[0..64] + "..." if str.length > 32
  str
end

puts "PstViewer 0.1.1a"

if ARGV.size < 2 then Process.exit end

File.open(ARGV.first, "rb") do |file|
  pstfile = Pst::PstFile.new(file)  
  message = pstfile.find_message(ARGV[1].hex)  
 
  puts "#<Pst::Message message_id=%08x>" % message.id
  message.properties.each do | property |
    puts "  %32s %-16s\t%s" % [ property.mapi_tag, '('+property.mapi_type+')', to_hex(property.value, property.mapi_type) ]
  end

  note = Pst::MapiMessage.create(message)
  puts note.to_rfc822
  
end
