# frozen_string_literal: true

module Writexlsx
  class Worksheet
    # Autofilter-related operations extracted from Worksheet to slim the main class.
    module Autofilter
      include Utility

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

      def prepare_filter_column(col) # :nodoc:
        # Check for a column reference in A1 notation and substitute.
        if col.to_s =~ /^\D/
          col_letter = col

          # Convert col ref to a cell ref and then to a col number.
          _dummy, col = substitute_cellref("#{col}1")
          raise "Invalid column '#{col_letter}'" if col >= COL_MAX
        end

        col_first, col_last = @filter_range

        # Reject column if it is outside filter range.
        raise "Column '#{col}' outside autofilter column range (#{col_first} .. #{col_last})" if col < col_first || col > col_last

        col
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

      private

      #
      # Extract the tokens from the filter expression. The tokens are mainly non-
      # whitespace groups. The only tricky part is to extract string tokens that
      # contain whitespace and/or quoted double quotes (Excel's escaped quotes).
      #
      def extract_filter_tokens(expression = nil) # :nodoc:
        return [] unless expression

        tokens = []
        str = expression
        while str =~ /"(?:[^"]|"")*"|\S+/
          tokens << ::Regexp.last_match(0)
          str = $LAST_MATCH_INFO.post_match
        end

        # Remove leading and trailing quotes and unescape other quotes
        tokens.map! do |token|
          token.sub!(/^"/, '')
          token.sub!(/"$/, '')
          token.gsub!('""', '"')

          # if token is number, convert to numeric.
          if token =~ /^([+-]?)(?=\d|\.\d)\d*(\.\d*)?([Ee]([+-]?\d+))?$/
            token.to_f == token.to_i ? token.to_i : token.to_f
          else
            token
          end
        end

        tokens
      end

      #
      # Converts the tokens of a possibly conditional expression into 1 or 2
      # sub expressions for further parsing.
      #
      def parse_filter_expression(expression, tokens) # :nodoc:
        # The number of tokens will be either 3 (for 1 expression)
        # or 7 (for 2  expressions).
        #
        if tokens.size == 7
          conditional = tokens[3]
          if conditional =~ /^(and|&&)$/
            conditional = 0
          elsif conditional =~ /^(or|\|\|)$/
            conditional = 1
          else
            raise "Token '#{conditional}' is not a valid conditional " \
                  "in filter expression '#{expression}'"
          end
          expression_1 = parse_filter_tokens(expression, tokens[0..2])
          expression_2 = parse_filter_tokens(expression, tokens[4..6])
          [expression_1, conditional, expression_2].flatten
        else
          parse_filter_tokens(expression, tokens)
        end
      end

      #
      # Parse the 3 tokens of a filter expression and return the operator and token.
      #
      def parse_filter_tokens(expression, tokens)     # :nodoc:
        operators = {
          '==' => 2,
          '='  => 2,
          '=~' => 2,
          'eq' => 2,

          '!=' => 5,
          '!~' => 5,
          'ne' => 5,
          '<>' => 5,

          '<'  => 1,
          '<=' => 3,
          '>'  => 4,
          '>=' => 6
        }

        operator = operators[tokens[1]]
        token    = tokens[2]

        # Special handling of "Top" filter expressions.
        if tokens[0] =~ /^top|bottom$/i
          value = tokens[1]
          if value.to_s =~ /\D/ || value.to_i < 1 || value.to_i > 500
            raise "The value '#{value}' in expression '#{expression}' " \
                  "must be in the range 1 to 500"
          end
          token.downcase!
          if token != 'items' && token != '%'
            raise "The type '#{token}' in expression '#{expression}' " \
                  "must be either 'items' or '%'"
          end

          operator = if tokens[0] =~ /^top$/i
                       30
                     else
                       32
                     end

          operator += 1 if tokens[2] == '%'

          token    = value
        end

        if !operator && tokens[0]
          raise "Token '#{tokens[1]}' is not a valid operator " \
                "in filter expression '#{expression}'"
        end

        # Special handling for Blanks/NonBlanks.
        if token.to_s =~ /^blanks|nonblanks$/i
          # Only allow Equals or NotEqual in this context.
          if operator != 2 && operator != 5
            raise "The operator '#{tokens[1]}' in expression '#{expression}' " \
                  "is not valid in relation to Blanks/NonBlanks'"
          end

          token.downcase!

          # The operator should always be 2 (=) to flag a "simple" equality in
          # the binary record. Therefore we convert <> to =.
          if token == 'blanks'
            token = ' ' if operator == 5
          elsif operator == 5
            operator = 2
            token    = 'blanks'
          else
            operator = 5
            token    = ' '
          end
        end

        # if the string token contains an Excel match character then change the
        # operator type to indicate a non "simple" equality.
        operator = 22 if operator == 2 && token.to_s =~ /[*?]/

        [operator, token]
      end
    end
  end
end
