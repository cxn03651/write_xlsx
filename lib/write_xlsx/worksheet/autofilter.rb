# frozen_string_literal: true

module Writexlsx
  class Worksheet
    # Autofilter-related operations extracted from Worksheet to slim the main class.
    module Autofilter
      #
      # :call-seq:
      #   autofilter(first_row, first_col, last_row, last_col)
      #
      # Set the autofilter area in the worksheet.
      #
      def autofilter(row1, col1 = nil, row2 = nil, col2 = nil)
        if (row_col_array = row_col_notation(row1))
          _row1, _col1, _row2, _col2 = row_col_array
        else
          _row1 = row1
          _col1 = col1
          _row2 = row2
          _col2 = col2
        end
        return if [_row1, _col1, _row2, _col2].include?(nil)

        # Reverse max and min values if necessary.
        _row1, _row2 = _row2, _row1 if _row2 < _row1
        _col1, _col2 = _col2, _col1 if _col2 < _col1

        @autofilter_area = convert_name_area(_row1, _col1, _row2, _col2)
        @autofilter_ref  = xl_range(_row1, _row2, _col1, _col2)
        @filter_range    = [_col1, _col2]

        # Store the filter cell positions for use in the autofit calculation.
        (_col1.._col2).each do |col|
          @filter_cells["#{_row1}:#{col}"] = 1
        end
      end

      #
      # Set the column filter criteria.
      #
      # The filter_column method can be used to filter columns in a autofilter
      # range based on simple conditions.
      #
      def filter_column(col, expression)
        raise "Must call autofilter before filter_column" unless @autofilter_area

        col = prepare_filter_column(col)

        tokens = extract_filter_tokens(expression)

        raise "Incorrect number of tokens in expression '#{expression}'" unless [3, 7].include?(tokens.size)

        tokens = parse_filter_expression(expression, tokens)

        # Excel handles single or double custom filters as default filters. We need
        # to check for them and handle them accordingly.
        if tokens.size == 2 && tokens[0] == 2
          # Single equality.
          filter_column_list(col, tokens[1])
        elsif tokens.size == 5 && tokens[0] == 2 && tokens[2] == 1 && tokens[3] == 2
          # Double equality with "or" operator.
          filter_column_list(col, tokens[1], tokens[4])
        else
          # Non default custom filter.
          @filter_cols[col] = Array.new(tokens)
          @filter_type[col] = 0
        end

        @filter_on = 1
      end

      #
      # Set the column filter criteria in Excel 2007 list style.
      #
      def filter_column_list(col, *tokens)
        tokens.flatten!
        raise "Incorrect number of arguments to filter_column_list" if tokens.empty?
        raise "Must call autofilter before filter_column_list" unless @autofilter_area

        col = prepare_filter_column(col)

        @filter_cols[col] = tokens
        @filter_type[col] = 1           # Default style.
        @filter_on        = 1
      end

      #
      # Write the <autoFilter> element.
      #
      def write_auto_filter # :nodoc:
        return unless autofilter_ref?

        attributes = [
          ['ref', @autofilter_ref]
        ]

        if filter_on?
          # Autofilter defined active filters.
          @writer.tag_elements('autoFilter', attributes) do
            write_autofilters
          end
        else
          # Autofilter defined without active filters.
          @writer.empty_tag('autoFilter', attributes)
        end
      end

      #
      # Function to iterate through the columns that form part of an autofilter
      # range and write the appropriate filters.
      #
      def write_autofilters # :nodoc:
        col1, col2 = @filter_range

        (col1..col2).each do |col|
          # Skip if column doesn't have an active filter.
          next unless @filter_cols[col]

          # Retrieve the filter tokens and write the autofilter records.
          tokens = @filter_cols[col]
          type   = @filter_type[col]

          # Filters are relative to first column in the autofilter.
          write_filter_column(col - col1, type, *tokens)
        end
      end

      #
      # Write the <filterColumn> element.
      #
      def write_filter_column(col_id, type, *filters) # :nodoc:
        @writer.tag_elements('filterColumn', [['colId', col_id]]) do
          if type == 1
            # Type == 1 is the new XLSX style filter.
            write_filters(*filters)
          else
            # Type == 0 is the classic "custom" filter.
            write_custom_filters(*filters)
          end
        end
      end

      #
      # Write the <filters> element.
      #
      def write_filters(*filters) # :nodoc:
        non_blanks = filters.reject { |filter| filter.to_s =~ /^blanks$/i }
        attributes = []

        attributes = [['blank', 1]] if filters != non_blanks

        if filters.size == 1 && non_blanks.empty?
          # Special case for blank cells only.
          @writer.empty_tag('filters', attributes)
        else
          # General case.
          @writer.tag_elements('filters', attributes) do
            non_blanks.sort.each { |filter| write_filter(filter) }
          end
        end
      end

      #
      # Write the <filter> element.
      #
      def write_filter(val) # :nodoc:
        @writer.empty_tag('filter', [['val', val]])
      end

      #
      # Write the <customFilters> element.
      #
      def write_custom_filters(*tokens) # :nodoc:
        if tokens.size == 2
          # One filter expression only.
          @writer.tag_elements('customFilters') { write_custom_filter(*tokens) }
        else
          # Two filter expressions.

          # Check if the "join" operand is "and" or "or".
          attributes = if tokens[2] == 0
                         [['and', 1]]
                       else
                         [['and', 0]]
                       end

          # Write the two custom filters.
          @writer.tag_elements('customFilters', attributes) do
            write_custom_filter(tokens[0], tokens[1])
            write_custom_filter(tokens[3], tokens[4])
          end
        end
      end

      #
      # Write the <customFilter> element.
      #
      def write_custom_filter(operator, val) # :nodoc:
        operators = {
          1  => 'lessThan',
          2  => 'equal',
          3  => 'lessThanOrEqual',
          4  => 'greaterThan',
          5  => 'notEqual',
          6  => 'greaterThanOrEqual',
          7  => 'startsWith',
          8  => 'notStartsWith',
          9  => 'endsWith',
          10 => 'notEndsWith',
          11 => 'contains',
          12 => 'notContains',
          13 => 'between',
          14 => 'notBetween'
        }

        attributes = [
          ['operator', operators[operator]],
          ['val', val]
        ]

        @writer.empty_tag('customFilter', attributes)
      end
    end
  end
end
