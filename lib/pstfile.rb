# pst.rb
# Copyright (c) 2009 David B. Conrad

$LOAD_PATH << '../lib'

require 'encryption'
require 'mapi_types'
require 'mapi_tags'
require 'iconv'
require 'message'

module Pst
  
  class PstFile

    ROOT_DESCRIPTOR_ID = 0x2
  
    attr_reader :io
  
    class FormatError < StandardError
    end

    def initialize io
      @io = io
      io.pos = 0
    
      @header = Header.new io.read(Header::SIZE)
      puts @header.inspect
        
      @message_store = MessageStore.new @io, @header.descriptor_idx_ptr 
      @file_store = FileStore.new @io, @header.data_struct_idx_ptr
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

    def read_block offset, size
      @io.seek offset
      buffer = @io.read size
      raise PstFile::FormatError, "tried to read #{size} bytes at 0x#{offset} but only got #{buffer.length}" if buffer.length != size
      buffer
    end

    def read_data_block data_id
      read_block find_block(data_id).file_offset, find_block(data_id).size
    end
    
    def read_enc_data_block data_id
      read_enc_block find_block(data_id).file_offset, find_block(data_id).size
    end
    
    def find_message id
      if @message_store.find id then
        messaage = Message.new self, id
      else
        # no message found with that id
      end
    end
    
    def find_message_data_block id
      @file_store.find @message_store.find(id).data_id
    end

    def find_message_assoc_block id
      assoc_data_id = @message_store.find(id).assoc_data_id
      @file_store.find assoc_data_id unless assoc_data_id == 0
    end
    
    def find_block id
      @file_store.find id
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

  class MessageStore

    BLOCK_SIZE = 512
  
    attr_reader :node_table
    attr_accessor :top_level_nodes, :root
  
    def initialize io, offset
      puts "Reading message store tree at 0x%08x..." % offset
      @node_table = Array.new
      self.validate!
    
      build_index io, offset
    
      puts "Added %d messages to the message store" % @node_table.size
    end

    def build_index io, offset
      read_node_recursive io, offset
    end
  
    def read_node_recursive io, offset
      io.seek offset
      node = IndexBlock.new io.read(MessageStore::BLOCK_SIZE)
    
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
      raise PstFile::FormatError, "no message with id 0x%08x in message store" % identifier   
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

  class FileStore

    BLOCK_SIZE = 512
  
    attr_reader :node_table
  
    def initialize io, offset
      puts "Reading file store tree at 0x%08x..." % offset
      @node_table = Array.new
      self.validate!
    
      build_index io, offset

      puts "Added %d nodes to the file store" % @node_table.size
    end

    def build_index io, offset
      read_node_recursive io, offset
    end
  
    def read_node_recursive io, offset
      io.seek offset
      node = IndexBlock.new io.read(FileStore::BLOCK_SIZE)
    
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
      raise PstFile::FormatError, "no data block with id 0x%08x in file store" % identifier   
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
            @node_table << MessageDescriptor.new(buffer)
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

  class MessageDescriptor

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
      "#<Pst::MessageDescriptor id=%08x, data_id=%08x, @assoc_data_id=%08x, @parent_id=%04x>" % [ @id, @data_id, @assoc_data_id, @parent_id ]
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

end