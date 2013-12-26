# -*- coding: utf-8 -*-

module Writexlsx
  class Chart
    class Chartline
      include Writexlsx::Utility

      attr_reader :line, :fill, :type

      def initialize(params)
        @line      = params[:line]
        @fill      = params[:fill]
        # Set the line properties for the marker..
        @line = line_properties(@line)
        # Allow 'border' as a synonym for 'line'.
        @line = line_properties(params[:border]) if params[:border]

        # Set the fill properties for the marker.
        @fill = fill_properties(@fill)
      end

      def line_defined?
        line && ptrue?(line[:_defined])
      end

      def fill_defined?
        fill && ptrue?(fill[:_defined])
      end
    end

    class Point < Chartline
    end

    class Gridline < Chartline
      attr_reader :visible

      def initialize(params)
        super(params)
        @visible = params[:visible]
      end
    end

    class Trendline < Chartline
      attr_reader :name, :forward, :backward, :order, :period

      def initialize(params)
        super(params)

        @name     = params[:name]
        @forward  = params[:forward]
        @backward = params[:backward]
        @order    = params[:order]
        @period   = params[:period]
        @type     = value_or_raise(types, params[:type], 'trendline type')
      end

      private

      def types
        {
          :exponential    => 'exp',
          :linear         => 'linear',
          :log            => 'log',
          :moving_average => 'movingAvg',
          :polynomial     => 'poly',
          :power          => 'power'
        }
      end
    end

    class Marker < Chartline
      attr_reader :size

      def initialize(params)
        super(params)

        if params[:type]
          @type = value_or_raise(types, params[:type], 'maker type')
        end

        @size      = params[:size]
        @automatic = false
        @automatic = true if @type == 'automatic'
      end

      def automatic?
        @automatic
      end

      private

      def types
        {
          :automatic  => 'automatic',
          :none       => 'none',
          :square     => 'square',
          :diamond    => 'diamond',
          :triangle   => 'triangle',
          :x          => 'x',
          :star       => 'star',
          :dot        => 'dot',
          :short_dash => 'dot',
          :dash       => 'dash',
          :long_dash  => 'dash',
          :circle     => 'circle',
          :plus       => 'plus',
          :picture    => 'picture'
        }
      end
    end

    class Errorbars
      include Writexlsx::Utility

      attr_reader :type, :direction, :endcap, :value, :line, :fill
      attr_reader :plus_values, :minus_values, :plus_data, :minus_data

      def initialize(params)
        @type = types[params[:type].to_sym] || 'fixedVal'

        # Set the value for error types that require it.
        @value = params[:value] || 1

        # Set the end-cap style.
        @endcap = params[:end_style] || 1

        # Set the error bar direction.
        case params[:direction]
        when 'minus'
          @direction = 'minus'
        when 'plus'
          @direction = 'plus'
        else
          @direction = 'both'
        end

        # Set any custom values
        @plus_values  = params[:plus_values]  || [1]
        @minus_values = params[:minus_values] || [1]
        @plus_data    = params[:plus_data]    || []
        @minus_data   = params[:minus_data]   || []

        # Set the line properties for the error bars.
        @line = line_properties(params[:line])
        @fill = params[:fill]
      end

      private

      def types
        {
          :fixed              => 'fixedVal',
          :percentage         => 'percentage',
          :standard_deviation => 'stdDev',
          :standard_error     => 'stdErr',
          :custom             => 'cust'
        }
      end
    end

    class Series
      include Writexlsx::Utility

      attr_reader :values, :categories, :name, :name_formula, :name_id
      attr_reader :cat_data_id, :val_data_id, :fill
      attr_reader :trendline, :smooth, :labels, :invert_if_neg
      attr_reader :x2_axis, :y2_axis, :error_bars, :points
      attr_accessor :line, :marker

      def initialize(chart, params = {})
        @values = aref_to_formula(params[:values])
        @categories = aref_to_formula(params[:categories])
        @name, @name_formula =
          chart.process_names(params[:name], params[:name_formula])
        @cat_data_id = chart.data_id(@categories, params[:categories_data])
        @val_data_id = chart.data_id(@values, params[:values_data])
        @name_id = chart.data_id(@name_formula, params[:name_data])
        if params[:border]
          @line = line_properties(params[:border])
        else
          @line = line_properties(params[:line])
        end
        @fill = fill_properties(params[:fill])
        @marker    = Marker.new(params[:marker]) if params[:marker]
        @trendline = Trendline.new(params[:trendline]) if params[:trendline]
        @smooth = params[:smooth]
        @error_bars = {
          :_x_error_bars => params[:x_error_bars] ? Errorbars.new(params[:x_error_bars]) : nil,
          :_y_error_bars => params[:y_error_bars] ? Errorbars.new(params[:y_error_bars]) : nil
        }
        @points = params[:points].collect { |p| p ? Point.new(p) : p } if params[:points]
        @labels = labels_properties(params[:data_labels])
        @invert_if_neg = params[:invert_if_negative]
        @x2_axis = params[:x2_axis]
        @y2_axis = params[:y2_axis]
      end

      def ==(other)
        methods = %w[categories values name name_formula name_id
                 cat_data_id val_data_id
                 line fill marker trendline
                 smooth labels invert_if_neg x2_axis y2_axis error_bars points ]
        methods.each do |method|
          return false unless self.instance_variable_get("@#{method}") == other.instance_variable_get("@#{method}")
        end
        true
      end

      def line_defined?
        line && ptrue?(line[:_defined])
      end

      private

      #
      # Convert and aref of row col values to a range formula.
      #
      def aref_to_formula(data) # :nodoc:
        # If it isn't an array ref it is probably a formula already.
        return data unless data.kind_of?(Array)
        xl_range_formula(*data)
      end

      #
      # Convert user defined labels properties to the structure required internally.
      #
      def labels_properties(labels) # :nodoc:
        return nil unless labels

        position = labels[:position]
        if position.nil? || position.empty?
          labels.delete(:position)
        else
          # Map user defined label positions to Excel positions.
          positions = {
            :center      => 'ctr',
            :right       => 'r',
            :left        => 'l',
            :top         => 't',
            :above       => 't',
            :bottom      => 'b',
            :below       => 'b',
            :inside_end  => 'inEnd',
            :outside_end => 'outEnd',
            :best_fit    => 'bestFit'
          }

          labels[:position] = value_or_raise(positions, position, 'label position')
        end

        labels
      end
    end
  end
end
