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

      def tag_elements(tag, attributes = [])
        start_tag(tag, attributes)
        yield
        end_tag(tag)
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
        str = "<#{tag}#{key_vals(attr)}/>"
        io_write(str)
      end

      def empty_tag_encoded(tag, attr = [])
        str = "<#{tag}#{key_vals(attr)}/>"
        io_write(str)
      end

      def data_element(tag, data, attr = [])
        tag_elements(tag, attr) { io_write("#{escape_data(data)}") }
      end

      #
      # Optimised tag writer ?  for shared strings <si> elements.
      #
      def si_element(data, attr)
        tag_elements('si') { data_element('t', data, attr) }
      end

      #
      # Optimised tag writer for shared strings <si> rich string elements.
      #
      def si_rich_element(data)
        io_write("<si>#{data}</si>")
      end

      def characters(data)
        io_write(escape_data(data))
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
          array << key_val(attr[i], escape_attributes(attr[i+1]))
        end
        array.join('')
      end

      def escape_attributes(str = '')
        return str if !(str =~ /["&<>]/)

        str.
          gsub(/&/, "&amp;").
          gsub(/"/, "&quot;").
          gsub(/</, "&lt;").
          gsub(/>/, "&gt;")
      end

      def escape_data(str = '')
        if str =~ /[&<>"]/
          str.gsub(/&/, '&amp;').
            gsub(/</, '&lt;').
            gsub(/>/, '&gt;')
        else
          str
        end
      end
    end
  end
end
