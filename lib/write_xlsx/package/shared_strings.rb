# -*- coding: utf-8 -*-
require 'write_xlsx/package/xml_writer_simple'
require 'write_xlsx/utility'

module Writexlsx
  module Package
    class SharedStrings

      include Writexlsx::Utility

      def initialize
        @writer        = Package::XMLWriterSimple.new
        @strings       = [] # string table
        @strings_index = {} # string table index
        @count         = {} # count
      end

      def index(string, params = {})
        add(string) unless params[:only_query]
        @strings_index[string]
      end

      def add(string)
        str = string.dup
        if @count[str]
          @count[str] += 1
        else
          @strings << str
          @strings_index[str] = @strings.size - 1
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
        write_xml_declaration do
          # Write the sst table.
          write_sst { write_sst_strings }
        end
      end

      private

      #
      # Write the <sst> element.
      #
      def write_sst
        schema       = 'http://schemas.openxmlformats.org'

        attributes =
          [
           ['xmlns',       schema + '/spreadsheetml/2006/main'],
           ['count',       total_count],
           ['uniqueCount', unique_count]
          ]

        @writer.tag_elements('sst', attributes) { yield }
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
        string = string.dup
        attributes = []

        # Excel escapes control characters with _xHHHH_ and also escapes any
        # literal strings of that type by encoding the leading underscore. So
        # "\0" -> _x0000_ and "_x0000_" -> _x005F_x0000_.
        # The following substitutions deal with those cases.

        # Escape the escape.
        string = string.gsub(/(_x[0-9a-fA-F]{4}_)/, '_x005F\1')

        # Convert control character to the _xHHHH_ escape.
        string = string.gsub(
                             /([\x00-\x08\x0B-\x1F])/,
                             sprintf("_x%04X_", $1.ord)
                             ) if string =~ /([\x00-\x08\x0B-\x1F])/

        # Convert character to \xC2\xxx or \xC3\xxx
        if string.bytesize == 1 && 0x80 <= string.ord && string.ord <= 0xFF
          string = add_c2_c3(string)
        end

        # Add attribute to preserve leading or trailing whitespace.
        attributes << ['xml:space', 'preserve'] if string =~ /\A\s|\s\Z/

        # Write any rich strings without further tags.
        if string =~ %r{^<r>} && string =~ %r{</r>$}
          @writer.si_rich_element(string)
        else
          @writer.si_element(string, attributes)
        end
      end

      def add_c2_c3(string)
        num = string.ord
        if 0x80 <= num && num < 0xC0
          0xC2.chr + num.chr
        else
          0xC3.chr + (num - 0x40).chr
        end
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
