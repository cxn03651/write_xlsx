# -*- coding: utf-8 -*-
# frozen_string_literal: true

require 'write_xlsx/constants'

module Writexlsx
  module Utility
    module Dimensions
      include Constants

      def check_dimensions(row, col)
        raise WriteXLSXDimensionError if !row || row >= ROW_MAX || !col || col >= COL_MAX

        0
      end

      #
      # Check that row and col are valid and store max and min values for use in
      # other methods/elements.
      #
      def check_dimensions_and_update_max_min_values(row, col, ignore_row = 0, ignore_col = 0)       # :nodoc:
        check_dimensions(row, col)
        store_row_max_min_values(row) if ignore_row == 0
        store_col_max_min_values(col) if ignore_col == 0

        0
      end

      def store_row_max_min_values(row)
        @dim_rowmin = row if !@dim_rowmin || (row < @dim_rowmin)
        @dim_rowmax = row if !@dim_rowmax || (row > @dim_rowmax)
      end

      def store_col_max_min_values(col)
        @dim_colmin = col if !@dim_colmin || (col < @dim_colmin)
        @dim_colmax = col if !@dim_colmax || (col > @dim_colmax)
      end
    end
  end
end
