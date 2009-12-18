# message.rb
# Copyright (c) 2009 David B. Conrad

$LOAD_PATH << '../lib'

require 'unpackle'

module Pst
  
  class Message
    attr_reader :id
    
    def initialize pst_file, message_id
      @pst_file = pst_file
      @id = message_id

      @data_block = @pst_file.find_message_data_block(@id)
      
      @assoc_data = AssociatedDataStore.new @pst_file, @pst_file.find_message_assoc_block(@id)
      @property_store = PropertyStore.new @pst_file, @data_block, @assoc_data
    end

    def message_class
      raise NotImplementedError, "Pst::Message::message_class not implemented"
      #@property_store.find_property_by_tag "PidTagMessageClass"
    end
    
    def properties
      @property_store
    end
    
    def attachments
      #TODO
      raise NotImplementedError, "Pst::Message::message_class not implemented"
    end
    
    def inspect
      str = "#<Pst::Message id=%08x data_ptr=%08x size=%08x>" % [ @id, @data_block.data_ptr, @data_block.size ]
      str << 0x0a << @assoc_data.inspect << 0x0a
      
      @property_store.each do | property |
        str << property.inspect << 0x0a
      end
      str
    end
  end

  class AssociatedDataStore
    
    SIGNATURE_OFFSET = 0x00
    NODE_TYPE_OFFSET = 0x01
    COUNT_OFFSET = 0x02
    TABLE_OFFSET = 0x04
  
    def initialize pst_file, assoc_data_block
      
      @pst_file = pst_file
      
      unless assoc_data_block.nil?
        @offset = assoc_data_block.data_ptr
        @block = @pst_file.read_block assoc_data_block.data_ptr, assoc_data_block.size
        @signature = @block[SIGNATURE_OFFSET, 1].unpack('C').first
        @node_type = @block[NODE_TYPE_OFFSET, 1].unpack('C').first
        @count = @block[COUNT_OFFSET, 2].unpack('v').first
        validate!
        build_data_chains
      end
      
    end
    
    def build_data_chains #TODO: add recursive build when next_block is not zero
      @node_table = Hash.new
      @block[TABLE_OFFSET..-1].scan(/(.{8})(.{8})(.{8})/m) do | entry |
        entry.map! { |e| e.unpackle('T').first }
        id, data_id, next_block = entry
        data_id &= 0xffffffff #TODO: high 32-bits of data_id is sometimes non-zero value. Why?
        raise NotImplementedError, 'multi-block associated data chain not supported' unless next_block == 0
        @node_table[id] = data_id
      end
    end
  
    def read id 
      buffer = @pst_file.read_data_block @node_table[id]
      data = String.new
      if buffer[0..1].unpack('v').first == 0x0101
        count, size = buffer[2..7].unpack('vV')
        buffer[8..-1].scan(/.{8}/m).each do | entry |
          segment_id = entry.unpack('Q').first
          data << @pst_file.read_enc_data_block(segment_id)
        end
        data
      end
    end
   
    def validate!
      raise PstFile::FormatError, 'unknown assoc data signature 0x%04x at offset %08x' % [ @signature, offset ] unless @signature == 0x02
    end
    
    def inspect
      str = "#<Pst::AssociatedDataStore offset=%08x sig=%02x node_type=%02x count=%04x [" % [ @offset, @signature, @node_type, @count ]
      unless @node_table.nil?
        @node_table.each do | id, data_id |
          str << " \"%08x => %08x\"" % [ id, data_id ]
        end
      end
      str << " ]>"
    end
    
  end

  class DataBlock
    include MAPI::Types

    INDEX_OFFSET = 0x00
    TABLE_TYPE_OFFSET = 0x02
    TABLE_HEADER_OFFSET = 0x04

    BLOCK_TYPES = { 0xbcec => 1 } # TODO: look at Outlook.pst/0x60d for table 7cec
    IMMEDIATE_TYPES = [
      PT_SHORT, PT_LONG, PT_BOOLEAN
    ]
  
    attr_reader :block, :size, :block_type, :signature, :element_size
  
    def initialize pst_file, data_block, assoc_data_block
      
      @assoc_data_block = assoc_data_block
      @block = pst_file.read_enc_block data_block.data_ptr, data_block.size
      @block_size = data_block.size
      @table_index_offset = @block[INDEX_OFFSET, 2].unpack('v').first
      @block_type = @block[TABLE_TYPE_OFFSET, 2].unpack('v').first
      @table_header_offset = @block[TABLE_HEADER_OFFSET, 4].unpack('V').first
      last_element_index = @block[@table_index_offset]
        
      raise PstFile::FormatError, 'unknown block type 0x%04x at offset %08x' % [ @block_type, data_block.size ] unless BLOCK_TYPES[@block_type]
      raise PstFile::FormatError, 'index offset 0x%04x beyond block of size 0x%04x' % [ @table_index_offset, @block_size ] unless @table_index_offset <= (@block_size-1)

      build_table_index
      
    end
  
    def build_table_index
      @table_index = Array.new
      offset = @table_index_offset + (@table_header_offset >> 4)
      entries = @block[offset..-1].unpack('v*') 
      entries.each_cons 2 do | ref_start, next_start |
        @table_index << [ ref_start, (next_start-1) ]
      end
    end
  
    def get_table_data offset
      case (offset & 0xf)
      when 0xf
        #TODO: data offset is an external reference
        buffer = @assoc_data_block.read offset
      else
        ref_start, ref_end = @table_index[(offset >> 4) / 2]
        buffer = @block[ref_start..ref_end]
      end
      buffer
    end

    def get_tuple_with_indirection property
     
      if IMMEDIATE_TYPES.include? property.type then
        # do nothing. immediate types do not need further processing
      else
        case property.type
        when MAPI::Types::PT_UNICODE
          ic = Iconv.new 'utf-8', 'utf-16le'
          property.value = ic.iconv(get_table_data property.value)
        else
          property.value = get_table_data property.value
        end                                           
      end
      property
    end
  
    def validate!
    end
  
  end

  class PropertyStore < DataBlock
    include Enumerable

    def initialize pst_file, data_block, assoc_data_block
      
      super pst_file, data_block, assoc_data_block
      
      table_header = get_table_data @table_header_offset
      @signature, @identifer_size, @value_size, @table_level, @descriptor_offset = table_header.unpack('CCCCV')
      
      build_property_array
      
      validate!
      
    end

    def build_property_array
      @properties = Array.new
      buffer = get_table_data(@descriptor_offset)
      buffer.scan(/.{8}/m) do | entry |
        key, type, value = entry.unpack('vvV')
        @properties << Property.new(key, type, value)
      end
    end
    
    def each
      @properties.each do | property |
        yield get_tuple_with_indirection property
      end
    end
  
    def validate!
      raise PstFile::FormatError, 'unknown table header signature 0x%04x' % @signature unless @signature == 0xb5
      raise PstFile::FormatError, 'unknown table identifer size 0x%04x' % @identifer_size unless @identifer_size == 0x02
      raise PstFile::FormatError, 'unknown table value size 0x%04x' % @value_size unless @value_size == 0x06
    end
  end

  class Property < Struct.new(:key, :type, :value)
    include MAPI::Types
    include MAPI::Tags
  
    def mapi_type
      type_str = PROPERTY_TYPES[self.type].first
      raise PstFile::FormatError, "Unknown PROPERTY_TYPE: type=%x" % self.type if type_str == nil
      type_str
    end
    
    def mapi_tag
      hash = PROPERTY_TAGS[ [self.key, self.mapi_type] ]
      #raise "Unknown PROPERTY_TAG: key=0x%04x type=0x%04x" % [ key, type ] if hash == nil
      if hash.nil?
        key_str = "UnknownType (0x%04x)" % self.key
      else
        key_str = hash[1]
        key_str = hash[0] if key_str.empty?
        key_str = "UnknownType (0x%04x)" % self.key if key_str.empty?
      end
      key_str
    end
    
    def inspect
      type_str = PROPERTY_TYPES[self.type].first
      raise PstFile::FormatError, "Unknown PROPERTY_TYPE: type=%x" % self.type if type_str == nil
      hash = PROPERTY_TAGS[ [self.key, type_str] ]
      #raise "Unknown PROPERTY_TAG: key=0x%04x type=0x%04x" % [ key, type ] if hash == nil
      if hash.nil?
        key_str = "UnknownType (0x%04x)" % self.key
      else
        key_str = hash[1]
        key_str = hash[0] if key_str.empty?
        key_str = "UnknownType (0x%04x)" % self.key if key_str.empty?
      end
      "#<Pst::Property %s (%s) value=%s>" % [ key_str, type_str, self.value.inspect ]
    end
  end

end
