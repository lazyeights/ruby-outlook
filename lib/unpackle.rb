# unpackle.rb - Extension to String class to support 8-byte little endian values
# Copyright (c) 2009 David B. Conrad

class String
    
    def unpackle format
      case format
      when 'T'
        array = self.unpack("V2")
        str = [ (array[0] << 32) + array[1] ]
      else    
        str = self.unpack(format)
      end
      str
    end
  
end
