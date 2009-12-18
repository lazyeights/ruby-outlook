# pstview - inspect Microsoft Outlook PST files
# Copyright (c) 2009 David B. Conrad

$LOAD_PATH << '../lib'

require 'pst_file'

puts "PstViewer 0.1.1a"

if ARGV.size < 2 then Process.exit end

File.open(ARGV.first, "rb") do |file|

  @pstfile = Pst::PstFile.new(file)  
  @message = @pstfile.find_message(ARGV[1].hex)
  #@message = @pstfile.root
  puts "#<Pst::Root message_id=%08x>" % @message.id
  @message.properties.each do | property |
    puts "  %s (%s) %s" % [ property.mapi_tag, property.mapi_type, property.value.inspect ]
  end

end
