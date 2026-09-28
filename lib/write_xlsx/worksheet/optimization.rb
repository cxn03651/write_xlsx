# -*- coding: utf-8 -*-
# frozen_string_literal: true

module Writexlsx
  class Worksheet
    #
    # Memory optimization mode, see Workbook#set_optimization. Each row is
    # written to a temporary file as soon as a later row is started, so the
    # cell table only holds the current row.
    #
    module Optimization
      def optimization? # :nodoc:
        @workbook.optimization?
      end

      private

      def setup_optimization
        @previous_row = 0
        return unless optimization?

        @cell_data_file = Tempfile.new(['write_xlsx', '.xml'])
        @cell_data_file.binmode
        @writer = Package::XMLWriterSimple.new(@cell_data_file)
      end

      #
      # In optimization mode, rows before the current one are already written
      # and can't be changed any more.
      #
      def row_already_written?(row)
        optimization? && row < @previous_row
      end

      #
      # Write out the previous row as a single row with cells and reset the
      # cell table. Spans are optional, so they aren't written in this mode.
      #
      def write_single_row(current_row = 0)
        row_num = @previous_row
        @previous_row = current_row

        return if not_contain_formatting_or_data?(row_num)

        if @cell_data_store[row_num]
          write_row_element(row_num, nil, *@set_rows[row_num]) do
            write_cell_column_dimension(row_num)
          end
        else
          # Row attributes or comments only.
          write_empty_row(row_num, nil, *@set_rows[row_num])
        end

        @cell_data_store = CellDataStore.new
      end

      #
      # Write the last row and switch from the temporary file to the writer of
      # the worksheet XML file.
      #
      def finish_optimized_rows
        write_single_row
        @writer = Package::XMLWriterSimple.new
      end

      #
      # Write the <sheetData> element from the rows in the temporary file.
      #
      def write_optimized_sheet_data
        @writer.tag_elements('sheetData') { @writer.append_file(@cell_data_file) }
      end

      def remove_optimized_rows
        @cell_data_file&.close!
        @cell_data_file = nil
      end
    end
  end
end
