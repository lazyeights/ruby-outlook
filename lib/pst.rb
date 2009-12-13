# pst.rb
# Copyright (c) 2009 David B. Conrad

$LOAD_PATH << '../lib'

require 'encryption'
require 'mapi_types'
require 'mapi_tags'
require 'iconv'

module Pst
  
  class PstFile

    ROOT_DESCRIPTOR_ID = 0x2
  
    attr_reader :io, :descriptors, :data_structures
  
    class FormatError < StandardError
    end

    def initialize io
      @io = io
      io.pos = 0
    
      @header = Header.new io.read(Header::SIZE)
      puts @header.inspect
        
      @descriptors = DescriptorIndex.new @io, @header.descriptor_idx_ptr 
      @data_structures = DataStructureIndex.new @io, @header.data_struct_idx_ptr
      #@extended_attributes = ExtendedAttributes.new @io. ...

    end
  
    def is_encrypted?
      @header.file_type == 0x17
    end

    def read_enc_block offset, size
      @io.seek offset
      buffer = @io.read size
      raise PstFile::FormatError, "tried to read #{size} bytes at 0x#{offset} but only got #{buffer.length}" if buffer.length != size
      if is_encrypted?
        CompressibleEncryption::decrypt buffer
      else
        buffer
      end
    end
  
  end

  class Item
    attr_reader :id
    
    def initialize pst_file, id
    
      @pst_file = pst_file
      @id = id
      descriptor = @pst_file.descriptors.find @id
      @data_struct = @pst_file.data_structures.find descriptor.data_id
      @property_store = PropertyStore.new @pst_file, @data_struct.file_offset, @data_struct.size
    
    end

    def to_s
      puts "#<Pst::Item descriptor=%08x data_struct=%08x>" % [ @id, @data_struct.file_offset ]
      @property_store.each do | property |
        puts property.inspect
      end
    end
  end

  class Header

    SIZE = 512
    MAGIC = 0x2142444e

    FILE_TYPE_OFFSET = 0x0a
    FILE_SIZE_OFFSET_64 = 0xb8
    DESCRIPTOR_IDX_BACK_PTR_64 = 0xd8
    DESCRIPTOR_IDX_OFFSET_64 = 0xe0  #index2
    DATA_STRUCT_IDX_BACK_PTR_64 = 0xe8
    DATA_STRUCT_IDX_OFFSET_64 = 0xf0 #index1
    IDX_ALLOC_FULL_MAP_64 = 0x180
    ENCRYPTION_OFFSET_64 = 0X201

    attr_reader :magic, :file_type, :file_size
    attr_reader :descriptor_idx_ptr, :descriptor_idx_back_ptr, :data_struct_idx_ptr, :data_struct_idx_back_ptr

    def initialize data
      puts "Reading header..."
      @magic = data.unpack('N').first
      @file_type = data[FILE_TYPE_OFFSET]
      @descriptor_idx_ptr = data[DESCRIPTOR_IDX_OFFSET_64, 8].unpack('Q').first
      @descriptor_idx_back_ptr = data[DESCRIPTOR_IDX_BACK_PTR_64, 8].unpack('Q').first
      @data_struct_idx_ptr = data[DATA_STRUCT_IDX_OFFSET_64, 8].unpack('Q').first
      @data_struct_idx_back_ptr = data[DATA_STRUCT_IDX_BACK_PTR_64, 8].unpack('Q').first
      @file_size = data[FILE_SIZE_OFFSET_64, 8].unpack('Q').first
      self.validate!
    end

    def inspect
      "#<Pst::Header descriptor b-tree=%08x, descriptor back ptr=%08x, block b-tree=%08x, block back ptr=%08x, file type=%04x, file size=%08x" % 
        [ @descriptor_idx_ptr, @descriptor_idx_back_ptr, @data_struct_idx_ptr, @data_struct_idx_back_ptr, @file_type, @file_size ]
    end
  
    def validate!    
      raise PstFile::FormatError, "bad signature on pst file (#{'0x%x' % @magic})" unless @magic == MAGIC 
    end
  
  end

  class DescriptorIndex

    BLOCK_SIZE = 512
  
    attr_reader :node_table
    attr_accessor :top_level_nodes, :root
  
    def initialize io, offset
      puts "Reading descriptor index tree at 0x%08x..." % offset
      @node_table = Array.new
      self.validate!
    
      build_index io, offset
    
      puts "Added %d nodes to the descriptor table" % @node_table.size
    end

    def build_index io, offset
      read_node_recursive io, offset
    end
  
    def read_node_recursive io, offset
      io.seek offset
      node = IndexBlock.new io.read(DescriptorIndex::BLOCK_SIZE)
    
      case node.node_level
      when 0  #this node contains leaf items
        node.node_table.each do | iterator |
          case iterator.parent_id
          when 0x0000
            @node_table << iterator
          when iterator.id
            @node_table << iterator
            @root = iterator
          else
            parent = self.find(iterator.parent_id)
            parent.children << iterator
            iterator.parent = parent
          end
        end
      else    #this node contains node items
        node.node_table.each do | iterator |
          read_node_recursive io, iterator.offset_ptr
        end
      end
    end
  
    def find identifier
      @node_table.each do | node |
        return node if node.id == identifier
        child = find_recursive identifier, node
        return child unless child == nil
      end
      raise PstFile::FormatError, "no descriptor with id 0x%08x in descriptor tree" % identifier   
    end
  
    def find_recursive identifier, parent
      parent.children.each do | node |
        return node if node.id == identifier
        child = find_recursive identifier, node
        return child unless child == nil      
      end
      nil
    end
  
    def validate!
    end
    
  end

  class DataStructureIndex

    BLOCK_SIZE = 512
  
    attr_reader :node_table
  
    def initialize io, offset
      puts "Reading data structure tree at 0x%08x..." % offset
      @node_table = Array.new
      self.validate!
    
      build_index io, offset

      puts "Added %d nodes to the data structure table" % @node_table.size
    end

    def build_index io, offset
      read_node_recursive io, offset
    end
  
    def read_node_recursive io, offset
      io.seek offset
      node = IndexBlock.new io.read(DataStructureIndex::BLOCK_SIZE)
    
      case node.node_level
      when 0  #this node contains leaf items
        @node_table += node.node_table
      else    #this node contains node items
        node.node_table.each do | iterator |
          read_node_recursive io, iterator.offset_ptr
        end
      end
    end

    def find identifier
      @node_table.each do | node |
        return node if node.id == identifier
      end
      raise PstFile::FormatError, "no data structure with id 0x%08x in data structure index" % identifier   
    end
  
    def validate!
    end
    
  end

  class IndexBlock
  
    BLOCK_SIZE = 512
    ITEM_COUNT_OFFSET_64 = 0x1e8
    ITEM_COUNT_MAX_OFFSET_64 = 0x1e9
    ITEM_SIZE_OFFSET_64 = 0x1ea
    NODE_LEVEL_OFFSET_64 = 0x1eb
    NODE_TYPE_OFFSET_64 = 0x1f0
    BACK_PTR_OFFSET_64 = 0X1f8
    ITEM_SIZE_64 = 24
    INDEX_COUNT_MAX_64 = 20
    DESC_COUNT_MAX_64 = 15
  
    attr_reader :block, :item_count, :item_size, :node_level, :node_table
  
    def initialize data
      @block = data
      @node_table = Array.new
      @item_count = @block[ITEM_COUNT_OFFSET_64]
      @item_size = @block[ITEM_SIZE_OFFSET_64]
      @item_max = @block[ITEM_COUNT_MAX_OFFSET_64]
      @node_level = @block[NODE_LEVEL_OFFSET_64]
      @node_type = @block[NODE_TYPE_OFFSET_64, 4].unpack('v').first
      self.validate!
    
      case @node_level
      when 0  #this node contains leaf items
        @block[0, @item_count * @item_size].scan(/.{#{@item_size}}/m).each do | buffer |
          case @node_type
          when 0x8181
            @node_table << Descriptor.new(buffer)
          when 0x8080
            @node_table << BlockPointer.new(buffer)
          end 
        end  
      else    #this node contains node items
        @block[0, @item_count * @item_size].scan(/.{#{@item_size}}/m).each do | buffer |
          @node_table << IndexBranchNode.new(buffer)
        end
      end

    end
  
    def validate!
      raise PstFile::FormatError, "unknown node type 0x%04x" % @node_type unless @node_type == 0x8080 || @node_type == 0x8181 
      raise PstFile::FormatError, "item max is unknown (%d)" % @item_max unless @item_max == INDEX_COUNT_MAX_64 || @item_max == DESC_COUNT_MAX_64
      raise PstFile::FormatError, "have too many active items in node (#{@item_count} items)" if @item_count > @item_max
    end
  
    def inspect
      "#<Pst::IndexBlock @node_type=%04x, @item_count=%02d, @item_size= %02d, @node_level=%02d>" [ @node_type  , @item_count, @item_size, @node_level ]
    end
  
  end

  class IndexBranchNode

    attr_reader :identifier, :parent_ptr, :offset_ptr

    def initialize data
      @id = data[0, 8].unpack('Q').first
      @parent_ptr = data[8, 8].unpack('Q').first
      @offset_ptr = data[16, 8].unpack('Q').first
      self.validate!
    end
  
    def validate!
      #raise PstFile::FormatError, "back_ptr 0x%08x in this node does not match required 0x%08x" % [ @parent_ptr, @parent_offset ] unless @parent_ptr == @parent_offset
    end

    def inspect
      "#<Pst::BranchNode id=%08x, @parent_ptr= %08x, @offset_ptr=%08x> " % [ @id, @parent_ptr, @offset_ptr ]
    end
  
  end

  class Descriptor

    attr_reader :id, :data_id, :assoc_data_id, :parent_id 
    attr_accessor :parent, :children

    def initialize data
      @id = data[0, 8].unpack('Q').first
      @data_id = data[8, 8].unpack('Q').first
      @assoc_data_id = data[16, 8].unpack('Q').first
      @parent_id = data[24, 4].unpack('V').first
    
      @parent = nil
      @children = Array.new
    
      self.validate!
    end
  
    def validate!
      puts self.inspect
      #raise PstFile::FormatError, "back_ptr 0x%08x in this node does not match required 0x%08x" % [ @parent_ptr, @parent_offset ] unless @parent_ptr == @parent_offset
    end
  
    def inspect
      "#<Pst::Descriptor id=%08x, data_id=%08x, @assoc_data_id=%08x, @parent_id=%04x>" % [ @id, @data_id, @assoc_data_id, @parent_id ]
    end
  
  end

  class BlockPointer

    attr_reader :id, :file_offset, :size, :alloc_ptr

    def initialize data
      @id = data[0, 8].unpack('Q').first
      @file_offset = data[8, 8].unpack('Q').first
      @size = data[16, 2].unpack('v').first
      @alloc_ptr = data[20, 4].unpack('V').first
      self.validate!
    end
  
    def validate!
      puts self.inspect
    end
  
    def inspect
      "#<Pst:BlockPointer id=%08x, file_offset=%08x, @size=%08x, @alloc_ptr= %08x>" % [ @id, @file_offset, @size, @alloc_ptr ]
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
  
    attr_reader :block, :size, :table_type, :signature, :element_size
  
    def initialize pst_file, offset, size

      @block = pst_file.read_enc_block offset, size
      @block_size = size
      @table_index_offset = @block[INDEX_OFFSET, 2].unpack('v').first
      @table_type = @block[TABLE_TYPE_OFFSET, 2].unpack('v').first
      @table_header_offset = @block[TABLE_HEADER_OFFSET, 4].unpack('V').first
      last_element_index = @block[@table_index_offset]
        
      build_table_index
    
      raise PstFile::FormatError, 'unknown table type 0x%04x' % @table_type unless BLOCK_TYPES[@table_type]
      raise PstFile::FormatError, 'index offset 0x%04x beyond table of size 0x%04x' % [ @table_index_offset, @block_size ] unless @table_index_offset <= (@block_size-1)

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
        raise PstFile::FormatError, 'unsupported table data offset %04x' % offset
      else
        ref_start, ref_end = @table_index[(offset >> 4) / 2]
        return @block[ref_start..ref_end]
      end
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

    def initialize pst_file, offset, size
      super
      table_header = get_table_data @table_header_offset
      @signature, @identifer_size, @value_size, @table_level, @descriptor_offset = table_header.unpack('CCCCV')

      @table_data = Array.new
      buffer = get_table_data(@descriptor_offset)
      buffer.scan(/.{8}/m) do | entry |
        key, type, value = entry.unpack('vvV')
        @table_data << Property.new(key, type, value)
      end
    
      validate!
    end
    ## TODO
    def each
      @table_data.each do | property |
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