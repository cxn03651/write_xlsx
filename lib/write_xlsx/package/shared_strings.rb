# -*- coding: utf-8 -*-
require 'write_xlsx/package/xml_writer_simple'
require 'write_xlsx/utility'

module Writexlsx
  module Package
    class SharedStrings

      include Writexlsx::Utility

      def initialize
        @writer = Package::XMLWriterSimple.new
        @strings      = []
        @string_count = 0
        @unique_count = 0
      end

      def set_xml_writer(filename)
        @writer.set_xml_writer(filename)
      end

      def assemble_xml_file
        write_xml_declaration

        # Write the sst table.
        write_sst(@string_count, @unique_count)

        # Write the sst strings.
        write_sst_strings

        # Close the sst tag.
        @writer.end_tag('sst')
        @writer.crlf
        @writer.close
      end

      #
      # Set the total sst string count.
      #
      def set_string_count(string_count)
        @string_count = string_count
      end

      #
      # Set the total of unique sst strings.
      #
      def set_unique_count(unique_count)
        @unique_count = unique_count
      end

      #
      # Add the array ref of strings to be written.
      #
      def add_strings(strings)
        @strings = strings
      end

      private

      def write_xml_declaration
        @writer.xml_decl
      end

      #
      # Write the <sst> element.
      #
      def write_sst(count, unique_count)
        schema       = 'http://schemas.openxmlformats.org'
        xmlns        = schema + '/spreadsheetml/2006/main'

        attributes = [
            'xmlns',       xmlns,
            'count',       count,
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
    end
  end
end
