# pstview - inspect Microsoft Outlook PST files
# Copyright (c) 2009 David B. Conrad

$LOAD_PATH << '../lib'

require 'pstfile'

puts "PstViewer 0.1.1a"

if ARGV.size < 2 then Process.exit end

File.open(ARGV.first, "rb") do |file|

  @pstfile = Pst::PstFile.new(file)  
  @item = @pstfile.find_message(ARGV[1].hex)
  puts @item
  
end
