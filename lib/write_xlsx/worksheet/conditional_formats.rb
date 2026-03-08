# frozen_string_literal: true

module Writexlsx
  class Worksheet
    # Conditional formatting operations extracted from Worksheet to slim the main class.
    module ConditionalFormats
      #
      # :call-seq:
      #   conditional_formatting(cell_or_cell_range, options)
      #
      # Conditional formatting is a feature of Excel which allows you to apply a
      # format to a cell or a range of cells based on a certain criteria.
      #
      def conditional_formatting(*args)
        cond_format = Package::ConditionalFormat.factory(self, *args)
        @cond_formats[cond_format.range] ||= []
        @cond_formats[cond_format.range] << cond_format
      end

      #
      # Write the <conditionalFormatting> element.
      #
      def write_conditional_formatting(range, cond_formats) # :nodoc:
        @writer.tag_elements('conditionalFormatting', [['sqref', range]]) do
          cond_formats.each(&:write_cf_rule)
        end
      end
    end
  end
end
