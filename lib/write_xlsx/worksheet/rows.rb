# frozen_string_literal: true

module Writexlsx
  class Worksheet
    # Row-related operations extracted from Worksheet to slim the main class.
    module Rows
      #
      # :call-seq:
      #   set_row(row [ , height, format, hidden, level, collapsed ])
      #
      # This method can be used to change the default properties of a row.
      # All parameters apart from +row+ are optional.
      #
      def set_row(*args)
        return unless args[0]

        row = args[0]
        height = args[1] || @default_height
        xf     = args[2]
        hidden = args[3] || 0
        level  = args[4] || 0
        collapsed = args[5] || 0

        # Use min col in check_dimensions. Default to 0 if undefined.
        min_col = @dim_colmin || 0

        # Check that row and col are valid and store max and min values.
        check_dimensions(row, min_col)
        store_row_col_max_min_values(row, min_col)

        height ||= @default_row_height

        # If the height is 0 the row is hidden and the height is the default.
        if height == 0
          hidden = 1
          height = @default_row_height
        end

        # Set the limits for the outline levels (0 <= x <= 7).
        level = 0 if level < 0
        level = 7 if level > 7

        @outline_row_level = level if level > @outline_row_level

        # Store the row properties.
        @set_rows[row] = [height, xf, hidden, level, collapsed]

        # Store the row change to allow optimisations.
        @row_size_changed = true

        # Store the row sizes for use when calculating image vertices.
        @row_sizes[row] = [height, hidden]
      end

      #
      # This method is used to set the height (in pixels) and the properties of the
      # row.
      #
      def set_row_pixels(*data)
        height = data[1]

        data[1] = pixels_to_height(height) if ptrue?(height)
        set_row(*data)
      end

      #
      # Set the default row properties
      #
      def set_default_row(height = nil, zero_height = nil)
        height      ||= @original_row_height
        zero_height ||= 0

        if height != @original_row_height
          @default_row_height = height

          # Store the row change to allow optimisations.
          @row_size_changed = 1
        end

        @default_row_zeroed = 1 if ptrue?(zero_height)
      end
    end
  end
end
