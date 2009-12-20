# storage_table.rb -  MAPI Storage Provider for Microsoft Outlook PST files
# Copyright (c) 2009 David B. Conrad

require 'pst_file'
require 'mapi_message'

module Pst

  class StorageTable
    
    def initialize filename
      @file = File.open(ARGV.first, "rb")
      @pst_file = Pst::PstFile.new(@file)
    end
    
    def find_message message_id
      message = MapiMessage.create(@pst_file.find_message message_id)
    end
    
  end
  
end
