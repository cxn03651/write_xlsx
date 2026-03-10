# -*- coding: utf-8 -*-
# frozen_string_literal: true

module Writexlsx
  class Worksheet
    module CellDataManager
      #
      # Store a CellData object in the worksheet cell table.
      #
      def store_data_to_table(cell_data, row, col) # :nodoc:
        @cell_data_store.store(cell_data, row, col)
      end

      #
      # Track row/col min/max used for range calculation.
      #
      def store_row_col_max_min_values(row, col)
        store_row_max_min_values(row)
        store_col_max_min_values(col)
      end

      #
      # Add a string to the shared string table, if it isn't already there, and
      # return the string index.
      #
      def shared_string_index(str) # :nodoc:
        @workbook.shared_string_index(str)
      end

      #
      # Returns a range of data from the worksheet cell table to be used in
      # chart cached data. Strings are returned as SST ids and decoded in the
      # workbook. Return nils for data that doesn't exist since Excel can chart
      # series with data missing.
      #
      def get_range_data(row_start, col_start, row_end, col_end)
        @cell_data_store.get_range_data(row_start, col_start, row_end, col_end)
      end

      private

      def cell_data_store # :nodoc:
        @cell_data_store
      end
    end
  end
end
