# pst_file.rb - File system interface for Microsoft Outlook PST files
# Copyright (c) 2009 David B. Conrad

$LOAD_PATH << '../../lib/pst'

require 'pst_header'
require 'pst_message'
require 'encryption'
require 'mapi_types'
require 'mapi_tags'
require 'iconv'

module Pst
  
  class PstFile

    ROOT_MESSAGE_ID = 0x21
    
    class FormatError < StandardError
    end

    def initialize io
      @io = io
      io.pos = 0
    
      @header = Header.new io.read(Header::SIZE)
      @message_store = MessageStore.new @io, @header.message_store_ptr 
      @file_store = DataStore.new @io, @header.data_store_ptr
      #@extended_attributes = ExtendedAttributes.new @io. ...

    end
  
    def root
      find_message(ROOT_MESSAGE_ID)
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

    def find_block data_id
      @file_store.find data_id
    end
    
    def read_data_block data_id
      read_block find_block(data_id).data_ptr, find_block(data_id).size
    end
    
    def read_enc_data_block data_id
      read_enc_block find_block(data_id).data_ptr, find_block(data_id).size
    end
    
    def find_message message_id
      if @message_store.find message_id then
        message = Message.new self, message_id
      else
        # no message found with that id
      end
    end
    
    def find_message_data_block message_id
      @file_store.find @message_store.find(message_id).data_id
    end

    def find_message_assoc_block message_id
      assoc_data_id = @message_store.find(message_id).assoc_data_id
      @file_store.find assoc_data_id unless assoc_data_id == 0
    end
    
  end

  class MessageStore

    BLOCK_SIZE = 512
  
    attr_reader :messages
    
    def initialize io, offset
      puts "Reading message store tree at 0x%08x..." % offset
      @messages = Array.new
      self.validate!
    
      build_index io, offset
    
      puts "Added %d messages to the message store" % @messages.size
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
            @messages << iterator
          when iterator.id
            @messages << iterator
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
  
    def find message_id
      @messages.each do | node |
        return node if node.id == message_id
        child = find_recursive message_id, node
        return child unless child == nil
      end
      raise PstFile::FormatError, "no message with id 0x%08x in message store" % message_id   
    end
  
    def find_recursive message_id, parent
      parent.children.each do | node |
        return node if node.id == message_id
        child = find_recursive message_id, node
        return child unless child == nil      
      end
      nil
    end
  
    def validate!
    end
    
  end

  class DataStore

    BLOCK_SIZE = 512
  
    def initialize io, offset
      puts "Reading file store tree at 0x%08x..." % offset
      @pointers = Array.new
      self.validate!
    
      build_index io, offset

      puts "Added %d nodes to the file store" % @pointers.size
    end

    def build_index io, offset
      read_node_recursive io, offset
    end
  
    def read_node_recursive io, offset
      io.seek offset
      node = IndexBlock.new io.read(DataStore::BLOCK_SIZE)
    
      case node.node_level
      when 0  #this node contains leaf items
        @pointers += node.node_table
      else    #this node contains node items
        node.node_table.each do | iterator |
          read_node_recursive io, iterator.offset_ptr
        end
      end
    end

    def find identifier
      @pointers.each do | node |
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
      @id, @parent_ptr, @offset_ptr = data[0..23].unpack('QQQ')
      self.validate!
    end
  
    def validate!
    end

    def inspect
      "#<Pst::BranchNode id=%08x, @parent_ptr= %08x, @offset_ptr=%08x> " % [ @id, @parent_ptr, @offset_ptr ]
    end
  
  end

  class MessageDescriptor

    attr_reader :id, :data_id, :assoc_data_id, :parent_id 
    attr_accessor :parent, :children

    def initialize data
      @id, @data_id, @assoc_data_id, @parent_id = data[0..27].unpack('QQQV')
    
      @parent = nil
      @children = Array.new
    
      self.validate!
    end
  
    def validate!
      #puts self.inspect
    end
  
    def inspect
      "#<Pst::MessageDescriptor id=%08x, data_id=%08x, @assoc_data_id=%08x, @parent_id=%04x>" % [ @id, @data_id, @assoc_data_id, @parent_id ]
    end
  
  end

  class BlockPointer

    attr_reader :id, :data_ptr, :size, :alloc_ptr

    def initialize data
      @id, @data_ptr, @size, @flag, @alloc_ptr = data[0..23].unpack('QQvvV')
      self.validate!
    end
  
    def validate!
      #puts self.inspect
    end
  
    def inspect
      "#<Pst:BlockPointer id=%08x, data_ptr=%08x, @size=%08x, @alloc_ptr= %08x>" % [ @id, @data_ptr, @size, @alloc_ptr ]
    end
  
  end

end