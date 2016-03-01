# -*- encoding: utf-8 -*-

module Writexlsx
  class Worksheet
    class CellData   # :nodoc:
      include Writexlsx::Utility

      attr_reader :row, :col, :token, :xf
      attr_reader :result, :range, :link_type, :url, :tip

      #
      # attributes for the <cell> element. This is the innermost loop so efficiency is
      # important where possible.
      #
      def cell_attributes #:nodoc:
        xf_index = xf ? xf.get_xf_index : 0
        attributes = [
                      ['r', xl_rowcol_to_cell(row, col)]
                     ]

        # Add the cell format index.
        if xf_index != 0
          attributes << ['s', xf_index]
        elsif @worksheet.set_rows[row] && @worksheet.set_rows[row][1]
          row_xf = @worksheet.set_rows[row][1]
          attributes << ['s', row_xf.get_xf_index]
        elsif @worksheet.col_formats[col]
          col_xf = @worksheet.col_formats[col]
          attributes << ['s', col_xf.get_xf_index]
        end
        attributes
      end

      def display_url_string?
        true
      end
    end

    class NumberCellData < CellData # :nodoc:
      def initialize(worksheet, row, col, num, xf)
        @worksheet = worksheet
        @row, @col, @token, @xf = row, col, num, xf
      end

      def data
        @token
      end

      def write_cell
        @worksheet.writer.tag_elements('c', cell_attributes) do
          @worksheet.write_cell_value(token)
        end
      end
    end

    class StringCellData < CellData # :nodoc:
      def initialize(worksheet, row, col, index, xf)
        @worksheet = worksheet
        @row, @col, @token, @xf = row, col, index, xf
      end

      def data
        { :sst_id => token }
      end

      def write_cell
        attributes = cell_attributes
        attributes << ['t', 's']
        @worksheet.writer.tag_elements('c', attributes) do
          @worksheet.write_cell_value(token)
        end
      end

      def display_url_string?
        false
      end
    end

    class FormulaCellData < CellData # :nodoc:
      def initialize(worksheet, row, col, formula, xf, result)
        @worksheet = worksheet
        @row, @col, @token, @xf, @result = row, col, formula, xf, result
      end

      def data
        @result || 0
      end

      def write_cell
        truefalse = {'TRUE' => 1, 'FALSE' => 0}
        error_code = ['#DIV/0!', '#N/A', '#NAME?', '#NULL!', '#NUM!', '#REF!', '#VALUE!']

        attributes = cell_attributes
        if @result &&  !(@result.to_s =~ /^([+-]?)(?=\d|\.\d)\d*(\.\d*)?([Ee]([+-]?\d+))?$/)
          if truefalse[@result]
            attributes << ['t', 'b']
            @result = truefalse[@result]
          elsif error_code.include?(@result)
            attributes << ['t', 'e']
          else
             attributes << ['t', 'str']
          end
        end
        @worksheet.writer.tag_elements('c', attributes) do
          @worksheet.write_cell_formula(token)
          @worksheet.write_cell_value(result || 0)
        end
      end
    end

    class FormulaArrayCellData < CellData # :nodoc:
      def initialize(worksheet, row, col, formula, xf, range, result)
        @worksheet = worksheet
        @row, @col, @token, @xf, @range, @result = row, col, formula, xf, range, result
      end

      def data
        @result || 0
      end

      def write_cell
        @worksheet.writer.tag_elements('c', cell_attributes) do
          @worksheet.write_cell_array_formula(token, range)
          @worksheet.write_cell_value(result)
        end
      end
    end

    class BlankCellData < CellData # :nodoc:
      def initialize(worksheet, row, col, xf)
        @worksheet = worksheet
        @row, @col, @xf = row, col, xf
      end

      def data
        ''
      end

      def write_cell
        @worksheet.writer.empty_tag('c', cell_attributes)
      end
    end
  end
end
