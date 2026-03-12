# -*- coding: utf-8 -*-
# frozen_string_literal: true

module Writexlsx
  module Utility
    module SheetnameQuoting
      #
      # Sheetnames used in references should be quoted if they contain any spaces,
      # special characters or if the look like something that isn't a sheet name.
      # TODO. We need to handle more special cases.
      #
      def quote_sheetname(sheetname) # :nodoc:
        name = sheetname.dup
        return name if already_quoted_sheetname?(name)
        return name unless sheetname_needs_quoting?(name)

        "'#{escape_sheetname(name)}'"
      end

      private

      def already_quoted_sheetname?(name)
        name.start_with?("'")
      end

      def escape_sheetname(name)
        name.gsub("'", "''")
      end

      def sheetname_needs_quoting?(name)
        contains_non_identifier_chars?(name) ||
          starts_with_digit_or_dot?(name) ||
          valid_a1_reference_name?(name) ||
          starts_with_rc_reference?(name) ||
          single_rc_reference?(name)
      end

      def contains_non_identifier_chars?(name)
        name.match?(/[^\p{L}\p{N}_.]/)
      end

      def starts_with_digit_or_dot?(name)
        name.match?(/^[\p{N}.]/)
      end

      def valid_a1_reference_name?(name)
        upcased = name.upcase
        return false unless upcased.match?(/^[A-Z]{1,3}\d+$/)

        row, col = xl_cell_to_rowcol(upcased)
        row.between?(0, 1_048_575) && col.between?(0, 16_383)
      end

      def starts_with_rc_reference?(name)
        upcased = name.upcase

        if (match = upcased.match(/^R(\d+)/))
          return match[1].to_i.between?(1, 1_048_576)
        end

        if (match = upcased.match(/^R?C(\d+)/))
          return match[1].to_i.between?(1, 16_384)
        end

        false
      end

      def single_rc_reference?(name)
        %w[R C RC].include?(name.upcase)
      end
    end
  end
end
