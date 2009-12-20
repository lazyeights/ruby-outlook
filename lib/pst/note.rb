# note.rb
# Copyright (c) 2009 David B. Conrad

require 'mapi_message'

module Pst

  class Note < MapiMessage
    
    attr_accessor :received_by, :delivered_to
    attr_accessor :from, :to, :cc, :bcc
    attr_accessor :sent_time
    attr_accessor :subject
    attr_accessor :body_plain, :body_html
    attr_accessor :attachments
    
    def to_rfc822
      str = "From: #{@from} \r\n"
      str += "To: " + @to + "\r\n"
      str += "Subject: #{@subject}\r\n"
      str += "Date: #{@sent_time}\r\n"
      str += "Content-Type: text/plain\r\n"
      str += "#{@body_plain}\r\n"
    end
    
  end

  class Attachment
  end
  
end