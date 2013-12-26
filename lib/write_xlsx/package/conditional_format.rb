# -*- coding: utf-8 -*-

module Writexlsx
  module Package
    class ConditionalFormat
      include Writexlsx::Utility

      def self.factory(worksheet, *args)
        range, param  =
          Package::ConditionalFormat.new(worksheet, nil, nil).
          range_param_for_conditional_formatting(*args)

        case param[:type]
        when 'cellIs'
          CellIsFormat.new(worksheet, range, param)
        when 'aboveAverage'
          AboveAverageFormat.new(worksheet, range, param)
        when 'top10'
          Top10Format.new(worksheet, range, param)
        when 'containsText', 'notContainsText', 'beginsWith', 'endsWith'
          TextOrWithFormat.new(worksheet, range, param)
        when 'timePeriod'
          TimePeriodFormat.new(worksheet, range, param)
        when 'containsBlanks', 'notContainsBlanks', 'containsErrors', 'notContainsErrors'
          BlanksOrErrorsFormat.new(worksheet, range, param)
        when 'colorScale'
          ColorScaleFormat.new(worksheet, range, param)
        when 'dataBar'
          DataBarFormat.new(worksheet, range, param)
        when 'expression'
          ExpressionFormat.new(worksheet, range, param)
        else # when 'duplicateValues', 'uniqueValues'
          ConditionalFormat.new(worksheet, range, param)
        end
      end

      attr_reader :range

      def initialize(worksheet, range, param)
        @worksheet, @range, @param = worksheet, range, param
        @writer = @worksheet.writer
      end

      def write_cf_rule
        @writer.empty_tag('cfRule', attributes)
      end

      def write_cf_rule_formula_tag(tag = formula)
        @writer.tag_elements('cfRule', attributes) do
          write_formula_tag(tag)
        end
      end

      def write_formula_tag(data) #:nodoc:
        data = data.sub(/^=/, '') if data.respond_to?(:sub)
        @writer.data_element('formula', data)
      end

      #
      # Write the <cfvo> element.
      #
      def write_cfvo(type, val)
        @writer.empty_tag('cfvo', [
                                   ['type', type],
                                   ['val', val]
                                  ])
      end

      def attributes
        attr = []
        attr << ['type' , type]
        attr << ['dxfId',    format]   if format
        attr << ['priority', priority]
        attr
      end

      def type
        @param[:type]
      end

      def format
        @param[:format]
      end

      def priority
        @param[:priority]
      end

      def criteria
        @param[:criteria]
      end

      def maximum
        @param[:maximum]
      end

      def minimum
        @param[:minimum]
      end

      def value
        @param[:value]
      end

      def direction
        @param[:direction]
      end

      def formula
        @param[:formula]
      end

      def min_type
        @param[:min_type]
      end
      def min_value
        @param[:min_value]
      end

      def min_color
        @param[:min_color]
      end

      def mid_type
        @param[:mid_type]
      end

      def mid_value
        @param[:mid_value]
      end

      def mid_color
        @param[:mid_color]
      end

      def max_type
        @param[:max_type]
      end

      def max_value
        @param[:max_value]
      end

      def max_color
        @param[:max_color]
      end

      def bar_color
        @param[:bar_color]
      end

      def range_param_for_conditional_formatting(*args)  # :nodoc:
        range_start_cell_for_conditional_formatting(*args)
        param_for_conditional_formatting(*args)

        handling_of_text_criteria        if @param[:type] == 'text'
        handling_of_time_period_criteria if @param[:type] == 'timePeriod'
        handling_of_blanks_error_types

        [@range, @param]
      end

      private

      def handling_of_text_criteria
        case @param[:criteria]
        when 'containsText'
          @param[:type]    = 'containsText';
          @param[:formula] =
            %Q!NOT(ISERROR(SEARCH("#{@param[:value]}",#{@start_cell})))!
        when 'notContains'
          @param[:type]    = 'notContainsText';
          @param[:formula] =
            %Q!ISERROR(SEARCH("#{@param[:value]}",#{@start_cell}))!
        when 'beginsWith'
          @param[:type] = 'beginsWith'
          @param[:formula] =
            %Q!LEFT(#{@start_cell},#{@param[:value].size})="#{@param[:value]}"!
        when 'endsWith'
          @param[:type] = 'endsWith'
          @param[:formula] =
            %Q!RIGHT(#{@start_cell},#{@param[:value].size})="#{@param[:value]}"!
        else
          raise "Invalid text criteria '#{@param[:criteria]} in conditional_formatting()"
        end
      end

      def handling_of_time_period_criteria
        case @param[:criteria]
        when 'yesterday'
          @param[:formula] = "FLOOR(#{@start_cell},1)=TODAY()-1"
        when 'today'
          @param[:formula] = "FLOOR(#{@start_cell},1)=TODAY()"
        when 'tomorrow'
          @param[:formula] = "FLOOR(#{@start_cell},1)=TODAY()+1"
        when 'last7Days'
          @param[:formula] =
            "AND(TODAY()-FLOOR(#{@start_cell},1)<=6,FLOOR(#{@start_cell},1)<=TODAY())"
        when 'lastWeek'
          @param[:formula] =
            "AND(TODAY()-ROUNDDOWN(#{@start_cell},0)>=(WEEKDAY(TODAY())),TODAY()-ROUNDDOWN(#{@start_cell},0)<(WEEKDAY(TODAY())+7))"
        when 'thisWeek'
          @param[:formula] =
            "AND(TODAY()-ROUNDDOWN(#{@start_cell},0)<=WEEKDAY(TODAY())-1,ROUNDDOWN(#{@start_cell},0)-TODAY()<=7-WEEKDAY(TODAY()))"
        when 'nextWeek'
          @param[:formula] =
            "AND(ROUNDDOWN(#{@start_cell},0)-TODAY()>(7-WEEKDAY(TODAY())),ROUNDDOWN(#{@start_cell},0)-TODAY()<(15-WEEKDAY(TODAY())))"
        when 'lastMonth'
          @param[:formula] =
            "AND(MONTH(#{@start_cell})=MONTH(TODAY())-1,OR(YEAR(#{@start_cell})=YEAR(TODAY()),AND(MONTH(#{@start_cell})=1,YEAR(A1)=YEAR(TODAY())-1)))"
        when 'thisMonth'
          @param[:formula] =
            "AND(MONTH(#{@start_cell})=MONTH(TODAY()),YEAR(#{@start_cell})=YEAR(TODAY()))"
        when 'nextMonth'
          @param[:formula] =
            "AND(MONTH(#{@start_cell})=MONTH(TODAY())+1,OR(YEAR(#{@start_cell})=YEAR(TODAY()),AND(MONTH(#{@start_cell})=12,YEAR(#{@start_cell})=YEAR(TODAY())+1)))"
        else
          raise "Invalid time_period criteria '#{@param[:criteria]}' in conditional_formatting()"
        end
      end

      def handling_of_blanks_error_types
        # Special handling of blanks/error types.
        case @param[:type]
        when 'containsBlanks'
          @param[:formula] = "LEN(TRIM(#{@start_cell}))=0"
        when 'notContainsBlanks'
          @param[:formula] = "LEN(TRIM(#{@start_cell}))>0"
        when 'containsErrors'
          @param[:formula] = "ISERROR(#{@start_cell})"
        when 'notContainsErrors'
          @param[:formula] = "NOT(ISERROR(#{@start_cell}))"
        when '2_color_scale'
          @param[:type] = 'colorScale'

          # Color scales don't use any additional formatting.
          @param[:format] = nil

          # Turn off 3 color parameters.
          @param[:mid_type]  = nil
          @param[:mid_color] = nil

          @param[:min_type]  ||= 'min'
          @param[:max_type]  ||= 'max'
          @param[:min_value] ||= 0
          @param[:max_value] ||= 0
          @param[:min_color] ||= '#FF7128'
          @param[:max_color] ||= '#FFEF9C'

          @param[:max_color] = palette_color( @param[:max_color] )
          @param[:min_color] = palette_color( @param[:min_color] )
        when '3_color_scale'
          @param[:type] = 'colorScale'

          # Color scales don't use any additional formatting.
          @param[:format] = nil

          @param[:min_type]  ||= 'min'
          @param[:mid_type]  ||= 'percentile'
          @param[:max_type]  ||= 'max'
          @param[:min_value] ||= 0
          @param[:mid_value] ||= 50
          @param[:max_value] ||= 0
          @param[:min_color] ||= '#F8696B'
          @param[:mid_color] ||= '#FFEB84'
          @param[:max_color] ||= '#63BE7B'

          @param[:max_color] = palette_color(@param[:max_color])
          @param[:mid_color] = palette_color(@param[:mid_color])
          @param[:min_color] = palette_color(@param[:min_color])
        when 'dataBar'
          # Color scales don't use any additional formatting.
          @param[:format] = nil

          @param[:min_type]  ||= 'min'
          @param[:max_type]  ||= 'max'
          @param[:min_value] ||= 0
          @param[:max_value] ||= 0
          @param[:bar_color] ||= '#638EC6'

          @param[:bar_color] = palette_color(@param[:bar_color])
        end
      end

      def palette_color(index)
        @worksheet.palette_color(index)
      end

      def range_start_cell_for_conditional_formatting(*args)  # :nodoc:
        row1, row2, col1, col2, user_range, param =
          row_col_param_for_conditional_formatting(*args)
        # If the first and last cell are the same write a single cell.
        if row1 == row2 && col1 == col2
          range = xl_rowcol_to_cell(row1, col1)
          @start_cell = range
        else
          range = xl_range(row1, row2, col1, col2)
          @start_cell = xl_rowcol_to_cell(row1, col1)
        end

        # Override with user defined multiple range if provided.
        range = user_range if user_range

        @range = range
      end

      def row_col_param_for_conditional_formatting(*args)
        # Check for a cell reference in A1 notation and substitute row and column
        if args[0] =~ /^\D/
          # Check for a user defined multiple range like B3:K6,B8:K11.
          user_range = args[0].sub(/^=/, '').gsub(/\s*,\s*/, ' ').gsub(/\$/, '') if args[0] =~ /,/
        end

        row1, col1, row2, col2, param = row_col_notation(args)
        if row2.respond_to?(:keys)
          param = row2
          row2, col2 = row1, col1
        end
        raise WriteXLSXInsufficientArgumentError if [row1, col1, row2, col2, param].include?(nil)

        # Check that row and col are valid without storing the values.
        check_dimensions(row1, col1)
        check_dimensions(row2, col2)

        # Swap last row/col for first row/col as necessary
        row1, row2 = row2, row1 if row1 > row2
        col1, col2 = col2, col1 if col1 > col2

        [row1, row2, col1, col2, user_range, param.dup]
      end

      def param_for_conditional_formatting(*args)  # :nodoc:
        dummy, dummy, dummy, dummy, dummy, @param =
          row_col_param_for_conditional_formatting(*args)
        check_conditional_formatting_parameters(@param)

        @param[:format] = @param[:format].get_dxf_index if @param[:format]
        @param[:priority] = @worksheet.dxf_priority
        @worksheet.dxf_priority += 1
      end

      def check_conditional_formatting_parameters(param)  # :nodoc:
        # Check for valid input parameters.
        unless (param.keys.uniq - valid_parameter_for_conditional_formatting).empty? &&
            param.has_key?(:type)                                   &&
            valid_type_for_conditional_formatting.has_key?(param[:type].downcase)
          raise WriteXLSXOptionParameterError, "Invalid type : #{param[:type]}"
        end

        param[:direction] = 'bottom' if param[:type] == 'bottom'
        param[:type] = valid_type_for_conditional_formatting[param[:type].downcase]

        # Check for valid criteria types.
        if param.has_key?(:criteria) && valid_criteria_type_for_conditional_formatting.has_key?(param[:criteria].downcase)
          param[:criteria] = valid_criteria_type_for_conditional_formatting[param[:criteria].downcase]
        end

        # Convert date/times value if required.
        if %w[date time cellIs].include?(param[:type])
          param[:type] = 'cellIs'

          param[:value]   = convert_date_time_if_required(param[:value])
          param[:minimum] = convert_date_time_if_required(param[:minimum])
          param[:maximum] = convert_date_time_if_required(param[:maximum])
        end

        # 'Between' and 'Not between' criteria require 2 values.
        if param[:criteria] == 'between' || param[:criteria] == 'notBetween'
          unless param.has_key?(:minimum) || param.has_key?(:maximum)
            raise WriteXLSXOptionParameterError, "Invalid criteria : #{param[:criteria]}"
          end
        else
          param[:minimum] = nil
          param[:maximum] = nil
        end

        # Convert date/times value if required.
        if param[:type] == 'date' || param[:type] == 'time'
          unless convert_date_time_value(param, :value) || convert_date_time_value(param, :maximum)
            raise WriteXLSXOptionParameterError
          end
        end
      end

      def convert_date_time_if_required(val)
        if val =~ /T/
          date_time = convert_date_time(val)
          raise "Invalid date/time value '#{val}' in conditional_formatting()" unless date_time
          date_time
        else
          val
        end
      end

      # List of valid input parameters for conditional_formatting.
      def valid_parameter_for_conditional_formatting
        [
         :type,
         :format,
         :criteria,
         :value,
         :minimum,
         :maximum,
         :min_type,
         :mid_type,
         :max_type,
         :min_value,
         :mid_value,
         :max_value,
         :min_color,
         :mid_color,
         :max_color,
         :bar_color
        ]
      end

      # List of  valid validation types for conditional_formatting.
      def valid_type_for_conditional_formatting
        {
          'cell'          => 'cellIs',
          'date'          => 'date',
          'time'          => 'time',
          'average'       => 'aboveAverage',
          'duplicate'     => 'duplicateValues',
          'unique'        => 'uniqueValues',
          'top'           => 'top10',
          'bottom'        => 'top10',
          'text'          => 'text',
          'time_period'   => 'timePeriod',
          'blanks'        => 'containsBlanks',
          'no_blanks'     => 'notContainsBlanks',
          'errors'        => 'containsErrors',
          'no_errors'     => 'notContainsErrors',
          '2_color_scale' => '2_color_scale',
          '3_color_scale' => '3_color_scale',
          'data_bar'      => 'dataBar',
          'formula'       => 'expression'
        }
      end

      # List of valid criteria types for conditional_formatting.
      def valid_criteria_type_for_conditional_formatting
        {
          'between'                  => 'between',
          'not between'              => 'notBetween',
          'equal to'                 => 'equal',
          '='                        => 'equal',
          '=='                       => 'equal',
          'not equal to'             => 'notEqual',
          '!='                       => 'notEqual',
          '<>'                       => 'notEqual',
          'greater than'             => 'greaterThan',
          '>'                        => 'greaterThan',
          'less than'                => 'lessThan',
          '<'                        => 'lessThan',
          'greater than or equal to' => 'greaterThanOrEqual',
          '>='                       => 'greaterThanOrEqual',
          'less than or equal to'    => 'lessThanOrEqual',
          '<='                       => 'lessThanOrEqual',
          'containing'               => 'containsText',
          'not containing'           => 'notContains',
          'begins with'              => 'beginsWith',
          'ends with'                => 'endsWith',
          'yesterday'                => 'yesterday',
          'today'                    => 'today',
          'last 7 days'              => 'last7Days',
          'last week'                => 'lastWeek',
          'this week'                => 'thisWeek',
          'next week'                => 'nextWeek',
          'last month'               => 'lastMonth',
          'this month'               => 'thisMonth',
          'next month'               => 'nextMonth'
        }
      end

      def date_1904?
        @worksheet.date_1904?
      end
    end

    class CellIsFormat < ConditionalFormat
      def attributes
        super << ['operator', criteria]
      end

      def write_cf_rule
        if minimum && maximum
          @writer.tag_elements('cfRule', attributes) do
            write_formula_tag(minimum)
            write_formula_tag(maximum)
          end
        else
          write_cf_rule_formula_tag(value)
        end
      end
    end

    class AboveAverageFormat < ConditionalFormat
      def attributes
        attr = super
        attr << ['aboveAverage', 0] if criteria =~ /below/
        attr << ['equalAverage', 1] if criteria =~ /equal/
        if criteria =~ /([123]) std dev/
          attr << ['stdDev', $~[1]]
        end
        attr
      end
    end

    class Top10Format < ConditionalFormat
      def attributes
        attr = super
        attr << ['percent', 1]             if criteria == '%'
        attr << ['bottom',  1]             if direction
        attr << ['rank',    (value || 10)]
        attr
      end
    end

    class TextOrWithFormat < ConditionalFormat
      def attributes
        attr = super
        attr << ['operator', criteria]
        attr << ['text',     value]
        attr
      end

      def write_cf_rule
        write_cf_rule_formula_tag
      end
    end

    class TimePeriodFormat < ConditionalFormat
      def attributes
        super << ['timePeriod', criteria]
      end

      def write_cf_rule
        write_cf_rule_formula_tag
      end
    end

    class BlanksOrErrorsFormat < ConditionalFormat
      def write_cf_rule
        write_cf_rule_formula_tag
      end
    end

    class ColorScaleFormat < ConditionalFormat
      def write_cf_rule
        @writer.tag_elements('cfRule', attributes) do
          write_color_scale
        end
      end

      #
      # Write the <colorScale> element.
      #
      def write_color_scale
        @writer.tag_elements('colorScale') do
          write_cfvo(min_type, min_value)
          write_cfvo(mid_type, mid_value) if mid_type
          write_cfvo(max_type, max_value)
          write_color(@writer, 'rgb', min_color)
          write_color(@writer, 'rgb', mid_color)  if mid_color
          write_color(@writer, 'rgb', max_color)
        end
      end
    end

    class DataBarFormat < ConditionalFormat
      def write_cf_rule
        @writer.tag_elements('cfRule', attributes) do
          write_data_bar
        end
      end

      #
      # Write the <dataBar> element.
      #
      def write_data_bar
        @writer.tag_elements('dataBar') do
          write_cfvo(min_type, min_value)
          write_cfvo(max_type, max_value)

          write_color(@writer, 'rgb', bar_color)
        end
      end
    end

    class ExpressionFormat < ConditionalFormat
      def write_cf_rule
        write_cf_rule_formula_tag(criteria)
      end
    end
  end
end
