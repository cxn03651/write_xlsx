# -*- coding: utf-8 -*-
# frozen_string_literal: true

require 'write_xlsx/utility/string_width'

module Writexlsx
  class Worksheet
    module RowColSizing
      include Writexlsx::Utility::StringWidth

      #
      # Convert the width of a cell from user's units to pixels. Excel rounds
      # the column width to the nearest pixel. If the width hasn't been set
      # by the user we use the default value. A hidden column is treated as
      # having a width of zero unless it has the special "object_position" of
      # 4 (size with cells).
      #
      def size_col(col, anchor = 0)
        info = col_info[col]
        calculate_col_pixels(
          info&.width,
          info&.hidden,
          anchor
        )
      end

      #
      # Convert the height of a cell from user's units to pixels. If the height
      # hasn't been set by the user we use the default value. A hidden row is
      # treated as having a height of zero unless it has the special
      # "object_position" of 4 (size with cells).
      #
      def size_row(row, anchor = 0)
        info = row_sizes[row]
        calculate_row_pixels(
          info&.first,
          info&.last,
          anchor
        )
      end

      private

      def calculate_col_pixels(width, hidden, anchor)
        width ||= @default_col_width

        return DEFAULT_COL_PIXELS unless width

        if hidden == 1 && anchor != 4
          0
        elsif width < 1
          ((width * (MAX_DIGIT_WIDTH + PADDING)) + 0.5).to_i
        else
          ((width * MAX_DIGIT_WIDTH) + 0.5).to_i + PADDING
        end
      end

      def calculate_row_pixels(height, hidden, anchor)
        height ||= default_row_height

        if hidden == 1 && anchor != 4
          0
        else
          (4 / 3.0 * height).to_i
        end
      end
    end
  end
end
