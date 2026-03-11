# frozen_string_literal: true

module Writexlsx
  class Worksheet
    module RichTextHelpers
      def cell_format_of_rich_string(rich_strings)
        # If the last arg is a format we use it as the cell format.
        rich_strings.pop if rich_strings[-1].respond_to?(:xf_index)
      end

      #
      # Convert the list of format, string tokens to pairs of (format, string)
      # except for the first string fragment which doesn't require a default
      # formatting run. Use the default for strings without a leading format.
      #
      def rich_strings_fragments(rich_strings) # :nodoc:
        # Create a temp format with the default font for unformatted fragments.
        default = Format.new(0)

        last = 'format'
        pos  = 0
        raw_string = ''

        fragments = []
        rich_strings.each do |token|
          if token.respond_to?(:xf_index)
            # Can't allow 2 formats in a row
            return nil if last == 'format' && pos > 0

            # Token is a format object. Add it to the fragment list.
            fragments << token
            last = 'format'
          else
            # Token is a string.
            if last == 'format'
              # If previous token was a format just add the string.
              fragments << token
            else
              # If previous token wasn't a format add one before the string.
              fragments << default << token
            end

            raw_string += token    # Keep track of actual string length.
            last = 'string'
          end
          pos += 1
        end
        [fragments, raw_string]
      end

      def xml_str_of_rich_string(fragments)
        # Create a temp XML::Writer object and use it to write the rich string
        # XML to a string.
        writer = Package::XMLWriterSimple.new

        # If the first token is a string start the <r> element.
        writer.start_tag('r') unless fragments[0].respond_to?(:xf_index)

        # Write the XML elements for the format string fragments.
        fragments.each do |token|
          if token.respond_to?(:xf_index)
            # Write the font run.
            writer.start_tag('r')
            token.write_font_rpr(writer, self)
          else
            # Write the string fragment part, with whitespace handling.
            attributes = []

            attributes << ['xml:space', 'preserve'] if token =~ /^\s/ || token =~ /\s$/
            writer.data_element('t', token, attributes)
            writer.end_tag('r')
          end
        end
        writer.string
      end
    end
  end
end
