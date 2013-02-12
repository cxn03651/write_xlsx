# -*- coding: utf-8 -*-

module Writexlsx
  module Package
    class ConditionalFormat
      include Writexlsx::Utility

      def self.factory(writer, range, param)
        case param[:type]
        when 'cellIs'
          CellIsFormat.new(writer, range, param)
        when 'aboveAverage'
          AboveAverageFormat.new(writer, range, param)
        when 'top10'
          Top10Format.new(writer, range, param)
        when 'containsText', 'notContainsText', 'beginsWith', 'endsWith'
          TextOrWithFormat.new(writer, range, param)
        when 'timePeriod'
          TimePeriodFormat.new(writer, range, param)
        when 'containsBlanks', 'notContainsBlanks', 'containsErrors', 'notContainsErrors'
          BlanksOrErrorsFormat.new(writer, range, param)
        when 'colorScale'
          ColorScaleFormat.new(writer, range, param)
        when 'dataBar'
          DataBarFormat.new(writer, range, param)
        when 'expression'
          ExpressionFormat.new(writer, range, param)
        else # when 'duplicateValues', 'uniqueValues'
          ConditionalFormat.new(writer, range, param)
        end
      end

      def initialize(writer, range, param)
        @writer, @range, @param = writer, range, param
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
        @writer.empty_tag('cfvo', ['type', type, 'val', val])
      end

      def attributes
        attr = ['type' , type]
        attr << 'dxfId'    << format   if format
        attr << 'priority' << priority
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
    end

    class CellIsFormat < ConditionalFormat
      def attributes
        super << 'operator' << criteria
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
        attr << 'aboveAverage' << 0 if criteria =~ /below/
        attr << 'equalAverage' << 1 if criteria =~ /equal/
        if criteria =~ /([123]) std dev/
          attr << 'stdDev'       << $~[1]
        end
        attr
      end
    end

    class Top10Format < ConditionalFormat
      def attributes
        attr = super
        attr << 'percent' << 1             if criteria == '%'
        attr << 'bottom'  << 1             if direction
        attr << 'rank'    << (value || 10)
        attr
      end
    end

    class TextOrWithFormat < ConditionalFormat
      def attributes
        attr = super
        attr << 'operator' << criteria
        attr << 'text'     << value
        attr
      end

      def write_cf_rule
        write_cf_rule_formula_tag
      end
    end

    class TimePeriodFormat < ConditionalFormat
      def attributes
        super << 'timePeriod' << criteria
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
