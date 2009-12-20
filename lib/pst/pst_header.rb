# pst_header.rb
# Copyright (c) 2009 David B. Conrad

$LOAD_PATH << '../../lib/pst'

module Pst
  
  class Header

    SIZE = 512
    MAGIC = 0x2142444e

    FILE_TYPE_OFFSET = 0x0a
    FILE_SIZE_OFFSET_64 = 0xb8
    MESSAGESTORE_OFFSET_64 = 0xe0
    DATASTORE_IDX_OFFSET_64 = 0xf0
    IDX_ALLOC_FULL_MAP_64 = 0x180
    ENCRYPTION_OFFSET_64 = 0X201

    attr_reader :magic, :file_type, :file_size
    attr_reader :message_store_ptr, :data_store_ptr

    def initialize data
      puts "Reading header..."
      @magic = data.unpack('N').first
      @file_type = data[FILE_TYPE_OFFSET]
      @message_store_ptr = data[MESSAGESTORE_OFFSET_64, 8].unpack('Q').first
      @data_store_ptr = data[DATASTORE_IDX_OFFSET_64, 8].unpack('Q').first
      @file_size = data[FILE_SIZE_OFFSET_64, 8].unpack('Q').first
      self.validate!
    end

    def inspect
      "#<Pst::Header descriptor b-tree=%08x, block b-tree=%08x, file type=%04x, file size=%08x" % 
        [ @message_store_ptr, @data_store_ptr, @file_type, @file_size ]
    end

    def validate!    
      raise PstFile::FormatError, "bad signature on pst file (#{'0x%x' % @magic})" unless @magic == MAGIC 
    end

  end
  
end
