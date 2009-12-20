# mapi_message.rb - Factory class for MAPI interface objects
# Copyright (c) 2009 David B. Conrad

module Pst

  class MapiMessage
    
    attr_accessor :properties
    
    def self.create pst_message
      case pst_message.message_class
      when 'IPM.Note'         # Message is an Email (note)
        msg = create_note pst_message
      when 'IPM.Contact'      # Message is a Contact
      when 'IPM.Appointment'  # Message is a Calendar entry
      when 'IPM.Activity'     # Message is a Journal entry
      when 'IPM.StickyNote'   # Message is a Notes entry
      when 'IPM.Task'         # Message is a Task
      else
        if (pst_message.id & 0x2)  # Message is a Folder
          msg = create_folder pst_message
        else
          raise NotImplementedError, "MAPI interface for message id %08x not supported" % pst_message.id
        end
      end
      msg
    end
    
    def self.create_note pst_message
      note = Note.new
      
      sender_type = pst_message.find_property("PidTagSenderAddressType").value
      name = pst_message.find_property("PidTagSenderName").value
      email = pst_message.find_property("PidTagSenderEmailAddress").value
      if sender_type == "SMTP"
        note.from = if name and email and name != email
            "#{name} <#{email}>"
          else
            email || name
          end
      else
        notes.from = name
      end
      
      note.to = pst_message.find_property("PidTagReceivedByName").value
      note.subject = pst_message.find_property("PidTagSubject").value
      note.sent_time = Time.at((pst_message.find_property("PidTagClientSubmitTime").value.unpack('Q').first - 116444736000000000) / 10000000)
      note.body_plain = pst_message.find_property("PidTagBody").value
      note.body_html = pst_message.find_property("PidTagHtml").value      
      note.properties = pst_message.properties
      note
    end
    
  end

end