# coding: utf-8
#
# XMLWriterSimple
#
require 'stringio'

module Writexlsx
  module Package
    class XMLWriterSimple
      XMLNS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

      def initialize
        @io = StringIO.new
      end

      def set_xml_writer(filename = nil)
        @filename = filename
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

      def tag_elements_str(tag, attributes = [])
        str = ''
        str << start_tag_str(tag, attributes)
        str << yield
        str << end_tag_str(tag)
      end

      def start_tag(tag, attr = [])
        io_write(start_tag_str(tag, attr))
      end

      def start_tag_str(tag, attr = [])
        "<#{tag}#{key_vals(attr)}>"
      end

      def end_tag(tag)
        io_write(end_tag_str(tag))
      end

      def end_tag_str(tag)
        "</#{tag}>"
      end

      def empty_tag(tag, attr = [])
        str = "<#{tag}#{key_vals(attr)}/>"
        io_write(str)
      end

      def empty_tag_encoded(tag, attr = [])
        io_write(empty_tag_encoded_str(tag, attr))
      end

      def empty_tag_encoded_str(tag, attr = [])
        "<#{tag}#{key_vals(attr)}/>"
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
        if @filename
          File.open(@filename, "wb:utf-8:utf-8") { |f| f << string }
        end
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

      def key_vals(attribute)
        attribute.
          inject('') { |str, attr| str + key_val(attr.first, escape_attributes(attr.last)) }
      end

      def escape_attributes(str = '')
        return str if !(str.to_s =~ /["&<>]/)

        str.
          gsub(/&/, "&amp;").
          gsub(/"/, "&quot;").
          gsub(/</, "&lt;").
          gsub(/>/, "&gt;")
      end

      def escape_data(str = '')
        if str.to_s =~ /[&<>]/
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
