# -*- coding: utf-8 -*-
require 'write_xlsx/package/xml_writer_simple'
require 'write_xlsx/utility'

module Writexlsx
  module Package
    class SharedStrings

      include Writexlsx::Utility

      def initialize
        @writer  = Package::XMLWriterSimple.new
        @strings = [] # string table
        @count   = {} # => count
      end

      def index(string)
        add(string)
        @strings.index(string)
      end

      def add(string)
        str = string.dup
        if @count[str]
          @count[str] += 1
        else
          @strings << str
          @count[str] = 1
        end
      end

      def string(index)
        @strings[index].dup
      end

      def empty?
        @strings.empty?
      end

      def set_xml_writer(filename)
        @writer.set_xml_writer(filename)
      end

      def assemble_xml_file
        write_xml_declaration

        # Write the sst table.
        write_sst

        # Write the sst strings.
        write_sst_strings

        # Close the sst tag.
        @writer.end_tag('sst')
        @writer.crlf
        @writer.close
      end

      private

      def write_xml_declaration
        @writer.xml_decl
      end

      #
      # Write the <sst> element.
      #
      def write_sst
        schema       = 'http://schemas.openxmlformats.org'

        attributes =
          [
           'xmlns',       schema + '/spreadsheetml/2006/main',
           'count',       total_count,
           'uniqueCount', unique_count
          ]

        @writer.start_tag('sst', attributes)
      end

      #
      # Write the sst string elements.
      #
      def write_sst_strings
        @strings.each { |string| write_si(string) }
      end

      #
      # Write the <si> element.
      #
      def write_si(string)
        attributes = []

        attributes << 'xml:space' << 'preserve' if string =~ /^[ \t]/ || string =~ /[ \t]$/

        @writer.start_tag('si')

        # Write any rich strings without further tags.
        if string =~ %r{^<r>} && string =~ %r{</r>$}
          @writer.io_write(string)
        else
          @writer.data_element('t', string, attributes)
        end

        @writer.end_tag('si')
      end

      def total_count
        @count.values.inject(0) { |sum, count| sum += count }
      end

      def unique_count
        @strings.size
      end
    end
  end
end
