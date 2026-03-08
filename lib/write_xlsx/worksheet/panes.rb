# frozen_string_literal: true

module Writexlsx
  class Worksheet
    # Pane-related operations extracted from Worksheet to slim the main class.
    module Panes
      ###############################################################################
      #
      # set_top_left_cell()
      #
      # Set the first visible cell at the top left of the worksheet.
      #
      def set_top_left_cell(row, col = nil)
        if (row_col_array = row_col_notation(row))
          _row, _col = row_col_array
        else
          _row = row
          _col = col
        end

        @top_left_cell = xl_rowcol_to_cell(_row, _col)
      end

      #
      # :call-seq:
      #   freeze_panes(row, col [ , top_row, left_col ] )
      #
      # This method can be used to divide a worksheet into horizontal or
      # vertical regions known as panes and to also "freeze" these panes so
      # that the splitter bars are not visible. This is the same as the
      # Window->Freeze Panes menu command in Excel
      #
      def freeze_panes(*args)
        return if args.empty?

        # Check for a cell reference in A1 notation and substitute row and column.
        if (row_col_array = row_col_notation(args.first))
          row, col, top_row, left_col = row_col_array
          type = args[1]
        else
          row, col, top_row, left_col, type = args
        end

        col      ||= 0
        top_row  ||= row
        left_col ||= col
        type     ||= 0

        @panes   = [row, col, top_row, left_col, type]
      end

      #
      # :call-seq:
      #   split_panes(y, x, top_row, left_col)
      #
      # Set panes and mark them as split.
      #
      def split_panes(*args)
        # Call freeze panes but add the type flag for split panes.
        freeze_panes(args[0], args[1], args[2], args[3], 2)
      end
    end
  end
end
