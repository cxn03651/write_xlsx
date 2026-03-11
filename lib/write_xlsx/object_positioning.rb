# -*- coding: utf-8 -*-
# frozen_string_literal: true

module Writexlsx
  module ObjectPositioning
    include Writexlsx::Utility

    #
    # Calculate the vertices that define the position of a graphical object
    # within the worksheet in pixels.
    #
    def position_object_pixels(col_start, row_start, x1, y1, width, height, anchor = nil) # :nodoc:
      context = object_positioning_context(anchor)

      position_object_pixels_with_context(
        col_start, row_start, x1, y1, width, height, context
      )
    end
    # def position_object_pixels(col_start, row_start, x1, y1, width, height, anchor = nil) # :nodoc:
    #   col_start, row_start, x1, y1 =
    #     adjust_start_position_for_negative_offsets(col_start, row_start, x1, y1)

    #   x_abs, y_abs =
    #     calculate_absolute_position(col_start, row_start, x1, y1, anchor)

    #   col_start, row_start, x1, y1 =
    #     adjust_start_position_for_cell_offsets(col_start, row_start, x1, y1, anchor)

    #   col_end, row_end, x2, y2 =
    #     calculate_object_end_position(col_start, row_start, x1, y1, width, height, anchor)

    #   [col_start, row_start, x1, y1, col_end, row_end, x2, y2, x_abs, y_abs]
    # end

    #
    # Calculate the vertices that define the position of a graphical object
    # within the worksheet in EMUs.
    #
    def position_object_emus(graphical_object) # :nodoc:
      object = graphical_object
      col_start, row_start, x1, y1, col_end, row_end, x2, y2, x_abs, y_abs =
        position_object_pixels(
          object.col, object.row, object.x_offset, object.y_offset,
          object.scaled_width, object.scaled_height, object.anchor
        )

      # Convert the pixel values to EMUs. See above.
      x1    = (0.5 + (9_525 * x1)).to_i
      y1    = (0.5 + (9_525 * y1)).to_i
      x2    = (0.5 + (9_525 * x2)).to_i
      y2    = (0.5 + (9_525 * y2)).to_i
      x_abs = (0.5 + (9_525 * x_abs)).to_i
      y_abs = (0.5 + (9_525 * y_abs)).to_i

      [col_start, row_start, x1, y1, col_end, row_end, x2, y2, x_abs, y_abs]
    end

    #
    # Convert the width of a cell from pixels to character units.
    #
    def pixels_to_width(pixels)
      max_digit_width = 7.0
      padding         = 5.0

      if pixels <= 12
        pixels / (max_digit_width + padding)
      else
        (pixels - padding) / max_digit_width
      end
    end

    #
    # Convert the height of a cell from pixels to character units.
    #
    def pixels_to_height(pixels)
      height = 0.75 * pixels
      height = height.to_i if (height - height.to_i).abs < 0.1
      height
    end

    private

    def position_object_pixels_with_context(col_start, row_start, x1, y1, width, height, context)
      col_start, row_start, x1, y1 =
        adjust_start_position_for_negative_offsets(col_start, row_start, x1, y1, context)

      x_abs, y_abs =
        calculate_absolute_position(col_start, row_start, x1, y1, context)

      col_start, row_start, x1, y1 =
        adjust_start_position_for_cell_offsets(col_start, row_start, x1, y1, context)

      col_end, row_end, x2, y2 =
        calculate_object_end_position(col_start, row_start, x1, y1, width, height, context)

      [col_start, row_start, x1, y1, col_end, row_end, x2, y2, x_abs, y_abs]
    end

    def adjust_start_position_for_negative_offsets(col_start, row_start, x1, y1, context)
      # Adjust start column for negative offsets.
      while x1 < 0 && col_start > 0
        x1 += context[:size_col].call(col_start - 1, 0)
        col_start -= 1
      end

      # Adjust start row for negative offsets.
      while y1 < 0 && row_start > 0
        y1 += context[:size_row].call(row_start - 1, 0)
        row_start -= 1
      end

      # Ensure that the image isn't shifted off the page at top left.
      x1 = 0 if x1 < 0
      y1 = 0 if y1 < 0

      [col_start, row_start, x1, y1]
    end

    def calculate_absolute_position(col_start, row_start, x1, y1, context)
      # Calculate the absolute x offset of the top-left vertex.
      x_abs = if context[:col_size_changed]
                (0..(col_start - 1)).inject(0) do |sum, col|
                  sum + context[:size_col].call(col, context[:anchor])
                end
              else
                # Optimisation for when the column widths haven't changed.
                context[:default_col_pixels] * col_start
              end
      x_abs += x1

      # Calculate the absolute y offset of the top-left vertex.
      y_abs = if context[:row_size_changed]
                (0..(row_start - 1)).inject(0) do |sum, row|
                  sum + context[:size_row].call(row, context[:anchor])
                end
              else
                # Optimisation for when the row heights haven't changed.
                context[:default_row_pixels] * row_start
              end
      y_abs += y1

      [x_abs, y_abs]
    end

    def adjust_start_position_for_cell_offsets(col_start, row_start, x1, y1, context)
      # Adjust start column for offsets that are greater than the col width.
      while x1 >= context[:size_col].call(col_start, context[:anchor])
        x1 -= context[:size_col].call(col_start, 0)
        col_start += 1
      end

      # Adjust start row for offsets that are greater than the row height.
      while y1 >= context[:size_row].call(row_start, context[:anchor])
        y1 -= context[:size_row].call(row_start, 0)
        row_start += 1
      end

      [col_start, row_start, x1, y1]
    end

    def calculate_object_end_position(col_start, row_start, x1, y1, width, height, context)
      # Initialise end cell to the same as the start cell.
      col_end = col_start
      row_end = row_start

      # Only offset the image in the cell if the row/col isn't hidden.
      width  += x1 if context[:size_col].call(col_start, context[:anchor]) > 0
      height += y1 if context[:size_row].call(row_start, context[:anchor]) > 0

      # Subtract the underlying cell widths to find the end cell of the object.
      while width >= context[:size_col].call(col_end, context[:anchor])
        width -= context[:size_col].call(col_end, context[:anchor])
        col_end += 1
      end

      # Subtract the underlying cell heights to find the end cell of the object.
      while height >= context[:size_row].call(row_end, context[:anchor])
        height -= context[:size_row].call(row_end, context[:anchor])
        row_end += 1
      end

      # The end vertices are whatever is left from the width and height.
      x2 = width
      y2 = height

      [col_end, row_end, x2, y2]
    end
  end
end
