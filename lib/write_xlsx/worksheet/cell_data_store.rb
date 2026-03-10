# -*- coding: utf-8 -*-
# frozen_string_literal: true

module Writexlsx
  class Worksheet
    class CellDataStore # :nodoc:
      def initialize
        @table = []
      end

      def [](row)
        @table[row]
      end

      def row(row)
        @table[row]
      end

      def row?(row)
        !@table[row].nil?
      end

      def store(cell_data, row, col)
        @table[row] ||= []
        @table[row][col] = cell_data
      end

      def fetch(row, col)
        return nil unless @table[row]

        @table[row][col]
      end

      def get_range_data(row_start, col_start, row_end, col_end)
        data = []

        (row_start..row_end).each do |row_num|
          unless row?(row_num)
            data << nil
            next
          end

          (col_start..col_end).each do |col_num|
            cell = fetch(row_num, col_num)
            data << (cell ? cell.data : nil)
          end
        end

        data
      end

      def each_row
        return enum_for(__method__) unless block_given?

        @table.each_with_index do |row_data, row_num|
          yield row_num, row_data if row_data
        end
      end
    end
  end
end
