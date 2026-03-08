# frozen_string_literal: true

module Writexlsx
  class Worksheet
    # Selection-related operations extracted from Worksheet to slim the main class.
    module Selection
      #
      # :call-seq:
      #   set_selection(cell_or_cell_range)
      #
      # Set which cell or cells are selected in a worksheet.
      #
      def set_selection(*args)
        return if args.empty?

        if (row_col_array = row_col_notation(args.first))
          row_first, col_first, row_last, col_last = row_col_array
        else
          row_first, col_first, row_last, col_last = args
        end

        active_cell = xl_rowcol_to_cell(row_first, col_first)

        if row_last  # Range selection.
          # Swap last row/col for first row/col as necessary
          row_first, row_last = row_last, row_first if row_first > row_last
          col_first, col_last = col_last, col_first if col_first > col_last

          sqref = xl_range(row_first, row_last, col_first, col_last)
        else          # Single cell selection.
          sqref = active_cell
        end

        # Selection isn't set for cell A1.
        return if sqref == 'A1'

        @selections = [[nil, active_cell, sqref]]
      end
    end
  end
end
