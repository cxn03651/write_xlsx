# -*- coding: utf-8 -*-
# frozen_string_literal: true

require 'write_xlsx/col_name'
require 'write_xlsx/constants'

module Writexlsx
  module Utility
    module CellReference
      include Constants

      #
      # xl_rowcol_to_cell($row, col, row_absolute, col_absolute)
      #
      def xl_rowcol_to_cell(row_or_name, col, row_absolute = false, col_absolute = false)
        if row_or_name.is_a?(Integer)
          row_or_name += 1      # Change from 0-indexed to 1 indexed.
        end
        col_str = xl_col_to_name(col, col_absolute)
        "#{col_str}#{absolute_char(row_absolute)}#{row_or_name}"
      end

      #
      # Returns: [row, col, row_absolute, col_absolute]
      #
      # The row_absolute and col_absolute parameters aren't documented because they
      # mainly used internally and aren't very useful to the user.
      #
      def xl_cell_to_rowcol(cell)
        cell =~ /(\$?)([A-Z]{1,3})(\$?)(\d+)/

        col_abs = ::Regexp.last_match(1) != ""
        col     = ::Regexp.last_match(2)
        row_abs = ::Regexp.last_match(3) != ""
        row     = ::Regexp.last_match(4).to_i

        # Convert base26 column string to number
        # All your Base are belong to us.
        chars = col.chars
        expn = 0
        col = 0

        chars.reverse.each do |char|
          col += (char.ord - 'A'.ord + 1) * (26**expn)
          expn += 1
        end

        # Convert 1-index to zero-index
        row -= 1
        col -= 1

        [row, col, row_abs, col_abs]
      end

      def xl_col_to_name(col, col_absolute)
        col_str = ColName.instance.col_str(col)
        if col_absolute
          "#{absolute_char(col_absolute)}#{col_str}"
        else
          # Do not allocate new string
          col_str
        end
      end

      def xl_range(row_1, row_2, col_1, col_2,
                   row_abs_1 = false, row_abs_2 = false, col_abs_1 = false, col_abs_2 = false)
        range1 = xl_rowcol_to_cell(row_1, col_1, row_abs_1, col_abs_1)
        range2 = xl_rowcol_to_cell(row_2, col_2, row_abs_2, col_abs_2)

        if range1 == range2
          range1
        else
          "#{range1}:#{range2}"
        end
      end

      def xl_range_formula(sheetname, row_1, row_2, col_1, col_2)
        # Use Excel's conventions and quote the sheet name if it contains any
        # non-word character or if it isn't already quoted.
        sheetname = quote_sheetname(sheetname)

        range1 = xl_rowcol_to_cell(row_1, col_1, 1, 1)
        range2 = xl_rowcol_to_cell(row_2, col_2, 1, 1)

        "=#{sheetname}!#{range1}:#{range2}"
      end

      # Check for a cell reference in A1 notation and substitute row and column
      def row_col_notation(row_or_a1)   # :nodoc:
        substitute_cellref(row_or_a1) if row_or_a1.respond_to?(:match) && row_or_a1.to_s =~ /^\D/
      end

      #
      # Substitute an Excel cell reference in A1 notation for  zero based row and
      # column values in an argument list.
      #
      # Ex: ("A4", "Hello") is converted to (3, 0, "Hello").
      #
      def substitute_cellref(cell, *args)       # :nodoc:
        normalized_cell = cell.upcase

        case normalized_cell
        # Convert a column range: 'A:A' or 'B:G'.
        # A range such as A:A is equivalent to A1:65536, so add rows as required
        when /\$?([A-Z]{1,3}):\$?([A-Z]{1,3})/
          row1, col1 =  xl_cell_to_rowcol(::Regexp.last_match(1) + '1')
          row2, col2 =  xl_cell_to_rowcol(::Regexp.last_match(2) + ROW_MAX.to_s)
          [row1, col1, row2, col2, *args]
        # Convert a cell range: 'A1:B7'
        when /\$?([A-Z]{1,3}\$?\d+):\$?([A-Z]{1,3}\$?\d+)/
          row1, col1 =  xl_cell_to_rowcol(::Regexp.last_match(1))
          row2, col2 =  xl_cell_to_rowcol(::Regexp.last_match(2))
          [row1, col1, row2, col2, *args]
        # Convert a cell reference: 'A1' or 'AD2000'
        when /\$?([A-Z]{1,3}\$?\d+)/
          row1, col1 = xl_cell_to_rowcol(::Regexp.last_match(1))
          [row1, col1, *args]
        else
          raise("Unknown cell reference #{normalized_cell}")
        end
      end
    end
  end
end
