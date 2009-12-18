# parse_mapitags.rb
# Copyright (c) 2009 David B. Conrad

require 'csv'
require '../lib/pst/mapi_types'

File.open("../data/MapiTags.csv", "r") do | input |
File.open("../data/MapiTags.out", "w") do | output |
  
  input.each do | line |     
    elements = CSV::parse_line(line)

    raise "Error on row: " + line if elements.size < 3

    while elements.size < 5 do
      elements << nil
    end
    
    unless elements[0].nil?
      output.write "  [ %s, '%s', '%s', '%s', '%s'],\n" % elements
    end
  end

end
end
