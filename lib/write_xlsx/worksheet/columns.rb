# frozen_string_literal: true

module Writexlsx
  class Worksheet
    # Column-related operations extracted from Worksheet to slim the main class.
    module Columns
      # :call-seq:
      #   set_column(firstcol, lastcol, width, format, hidden, level, collapsed)
      #
      # This method can be used to change the default properties of a single
      # column or a range of columns. All parameters apart from +first_col+
      # and +last_col+ are optional.
      def set_column(*args)
        # Check for a cell reference in A1 notation and substitute row and column
        # ruby 3.2 no longer handles =~ for various types
        if args[0].respond_to?(:=~) && args[0].to_s =~ /^\D/
          _row1, firstcol, _row2, lastcol, *data = substitute_cellref(*args)
        else
          firstcol, lastcol, *data = args
        end

        # Ensure at least firstcol, lastcol and width
        return unless firstcol && lastcol && !data.empty?

        # Assume second column is the same as first if 0. Avoids KB918419 bug.
        lastcol = firstcol unless ptrue?(lastcol)

        # Ensure 2nd col is larger than first. Also for KB918419 bug.
        firstcol, lastcol = lastcol, firstcol if firstcol > lastcol

        width, format, hidden, level, collapsed = data
        autofit = 0

        # Check that cols are valid and store max and min values with default row.
        # NOTE: The check shouldn't modify the row dimensions and should only modify
        #       the column dimensions in certain cases.
        ignore_row = 1
        ignore_col = 1
        ignore_col = 0 if format.respond_to?(:xf_index)   # Column has a format.
        ignore_col = 0 if width && ptrue?(hidden)         # Column has a width but is hidden

        check_dimensions_and_update_max_min_values(0, firstcol, ignore_row, ignore_col)
        check_dimensions_and_update_max_min_values(0, lastcol,  ignore_row, ignore_col)

        # Set the limits for the outline levels (0 <= x <= 7).
        level ||= 0
        level = 0 if level < 0
        level = 7 if level > 7

        # Excel has a maximum column width of 255 characters.
        width = 255.0 if width && width > 255.0

        @outline_col_level = level if level > @outline_col_level

        # Store the column data based on the first column. Padded for sorting.
        (firstcol..lastcol).each do |col|
          @col_info[col] =
            COLINFO.new(width, format, hidden, level, collapsed, autofit)
        end

        # Store the column change to allow optimisations.
        @col_size_changed = true
      end

      #
      # Set the width (and properties) of a single column or a range of columns in
      # pixels rather than character units.
      #
      def set_column_pixels(*data)
        cell = data[0]

        # Check for a cell reference in A1 notation and substitute row and column
        if cell =~ /^\D/
          data = substitute_cellref(*data)

          # Returned values row1 and row2 aren't required here. Remove them.
          data.shift         # $row1
          data.delete_at(1)  # $row2
        end

        # Ensure at least $first_col, $last_col and $width
        return if data.size < 3

        first_col, last_col, pixels, format, hidden, level = data
        hidden ||= 0

        width = pixels_to_width(pixels) if ptrue?(pixels)

        set_column(first_col, last_col, width, format, hidden, level)
      end

      #
      # autofit()
      #
      # Simulate autofit based on the data, and datatypes in each column. We do this
      # by estimating a pixel width for each cell data.
      #
      def autofit
        col_width = {}

        # Iterate through all the data in the worksheet.
        (@dim_rowmin..@dim_rowmax).each do |row_num|
          # Skip row if it doesn't contain cell data.
          next unless @cell_data_table[row_num]

          (@dim_colmin..@dim_colmax).each do |col_num|
            length = 0
            case (cell_data = @cell_data_table[row_num][col_num])
            when StringCellData, RichStringCellData
              # Handle strings and rich strings.
              #
              # For standard shared strings we do a reverse lookup
              # from the shared string id to the actual string. For
              # rich strings we use the unformatted string. We also
              # split multiline strings and handle each part
              # separately.
              string = cell_data.raw_string

              length = if string =~ /\n/
                         # Handle multiline strings.
                         max = string.split("\n").collect do |str|
                           xl_string_pixel_width(str)
                         end.max
                       else
                         xl_string_pixel_width(string)
                       end
            when DateTimeCellData

              # Handle dates.
              #
              # The following uses the default width for mm/dd/yyyy
              # dates. It isn't feasible to parse the number format
              # to get the actual string width for all format types.
              length = @default_date_pixels
            when NumberCellData

              # Handle numbers.
              #
              # We use a workaround/optimization for numbers since
              # digits all have a pixel width of 7. This gives a
              # slightly greater width for the decimal place and
              # minus sign but only by a few pixels and
              # over-estimation is okay.
              length = 7 * cell_data.token.to_s.length
            when BooleanCellData

              # Handle boolean values.
              #
              # Use the Excel standard widths for TRUE and FALSE.
              length = if ptrue?(cell_data.token)
                         31
                       else
                         36
                       end
            when FormulaCellData, FormulaArrayCellData, DynamicFormulaArrayCellData
              # Handle formulas.
              #
              # We only try to autofit a formula if it has a
              # non-zero value.
              if ptrue?(cell_data.data)
                length = xl_string_pixel_width(cell_data.data)
              end
            end

            # If the cell is in an autofilter header we add an
            # additional 16 pixels for the dropdown arrow.
            if length > 0 &&
               @filter_cells["#{row_num}:#{col_num}"]
              length += 16
            end

            # Add the string lenght to the lookup hash.
            max                = col_width[col_num] || 0
            col_width[col_num] = length if length > max
          end
        end

        # Apply the width to the column.
        col_width.each do |col_num, pixel_width|
          # Convert the string pixel width to a character width using an
          # additional padding of 7 pixels, like Excel.
          width = pixels_to_width(pixel_width + 7)

          # The max column character width in Excel is 255.
          width = 255.0 if width > 255.0

          # Add the width to an existing col info structure or add a new one.
          if @col_info[col_num]
            @col_info[col_num].width   = width
            @col_info[col_num].autofit = 1
          else
            @col_info[col_num] =
              COLINFO.new(width, nil, 0, 0, 0, 1)
          end
        end
      end
    end
  end
end
