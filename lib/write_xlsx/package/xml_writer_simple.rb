# coding: utf-8
#
# XMLWriterSimple
#
require 'stringio'

module Writexlsx
  module Package
    class XMLWriterSimple
      def initialize
        @io = StringIO.new
      end

      def set_xml_writer(filename = nil)
        fh = File.open(filename, "wb")

        @io = fh
      end

      def xml_decl(encoding = 'UTF-8', standalone = true)
        str = %Q!<?xml version="1.0" encoding="#{encoding}" standalone="#{standalone ? 'yes' : 'no'}"?>\n!
        io_write(str)
      end

      def start_tag(tag, attr = [])
        str = "<#{tag}#{key_vals(attr)}>"
        io_write(str)
      end

      def end_tag(tag)
        str = "</#{tag}>"
        io_write(str)
      end

      def empty_tag(tag, attr = [])
        str = "<#{tag}#{key_vals(attr)} />"
        io_write(str)
      end

      def data_element(tag, data, attr = [])
        str = start_tag(tag, attr)
        str << io_write("#{characters(data)}")
        str << end_tag(tag)
      end

      def characters(data)
        escape_xml_chars(data)
      end

      def crlf
        io_write("\n")
      end

      def close
        @io.close
      end

      def string
        @io.string
      end

      def io_write(str)
        @io << str
        str
      end

      private

      def key_val(key, val)
        %Q{ #{key}="#{val}"}
      end

      def key_vals(attr)
        array = []
        (0 .. attr.size-1).step(2) do |i|
          array << key_val(attr[i], escape_xml_chars(attr[i+1]))
        end
        array.join('')
      end

      def escape_xml_chars(str = '')
        if str =~ /[&<>"]/
          str.gsub(/&/, '&amp;').gsub(/</, '&lt;').gsub(/>/, '&gt;').gsub(/"/, '&quot;')
        else
          str
        end
      end
    end
  end
end
