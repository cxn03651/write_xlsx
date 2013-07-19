# -*- coding: utf-8 -*-

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
    @cat_data_id = chart.get_data_id(@categories, params[:categories_data])
    @val_data_id = chart.get_data_id(@values, params[:values_data])
    @name_id = chart.get_data_id(@name_formula, params[:name_data])
    if params[:border]
      @line = line_properties(params[:border])
    else
      @line = line_properties(params[:line])
    end
    @fill = fill_properties(params[:fill])
    @marker = marker_properties(params[:marker])
    @trendline = trendline_properties(params[:trendline])
    @smooth = params[:smooth]
    @error_bars = {
      :_x_error_bars => error_bars_properties(params[:x_error_bars]),
      :_y_error_bars => error_bars_properties(params[:y_error_bars])
    }
    @points = points_properties(params[:points])
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
  # Convert user defined line properties to the structure required internally.
  #
  def line_properties(line) # :nodoc:
    return { :_defined => 0 } unless line

    dash_types = {
      :solid               => 'solid',
      :round_dot           => 'sysDot',
      :square_dot          => 'sysDash',
      :dash                => 'dash',
      :dash_dot            => 'dashDot',
      :long_dash           => 'lgDash',
      :long_dash_dot       => 'lgDashDot',
      :long_dash_dot_dot   => 'lgDashDotDot',
      :dot                 => 'dot',
      :system_dash_dot     => 'sysDashDot',
      :system_dash_dot_dot => 'sysDashDotDot'
    }

    # Check the dash type.
    dash_type = line[:dash_type]

    if dash_type
      line[:dash_type] = value_or_raise(dash_types, dash_type, 'dash type')
    end

    line[:_defined] = 1

    line
  end

  #
  # Convert user defined fill properties to the structure required internally.
  #
  def fill_properties(fill) # :nodoc:
    return { :_defined => 0 } unless fill

    fill[:_defined] = 1

    fill
  end

  #
  # Convert user defined marker properties to the structure required internally.
  #
  def marker_properties(marker) # :nodoc:
    return unless marker

    types = {
      :automatic  => 'automatic',
      :none       => 'none',
      :square     => 'square',
      :diamond    => 'diamond',
      :triangle   => 'triangle',
      :x          => 'x',
      :star       => 'start',
      :dot        => 'dot',
      :short_dash => 'dot',
      :dash       => 'dash',
      :long_dash  => 'dash',
      :circle     => 'circle',
      :plus       => 'plus',
      :picture    => 'picture'
    }

    # Check for valid types.
    marker_type = marker[:type]

    if marker_type
      marker[:automatic] = 1 if marker_type == 'automatic'
      marker[:type] = value_or_raise(types, marker_type, 'maker type')
    end

    # Set the line properties for the marker..
    line = line_properties(marker[:line])

    # Allow 'border' as a synonym for 'line'.
    line = line_properties(marker[:border]) if marker[:border]

    # Set the fill properties for the marker.
    fill = fill_properties(marker[:fill])

    marker[:_line] = line
    marker[:_fill] = fill

    marker
  end

  #
  # Convert user defined trendline properties to the structure required internally.
  #
  def trendline_properties(trendline) # :nodoc:
    return unless trendline

    types = {
      :exponential    => 'exp',
      :linear         => 'linear',
      :log            => 'log',
      :moving_average => 'movingAvg',
      :polynomial     => 'poly',
      :power          => 'power'
    }

    # Check the trendline type.
    trend_type = trendline[:type]

    trendline[:type] = value_or_raise(types, trend_type, 'trendline type')

    # Set the line properties for the trendline..
    line = line_properties(trendline[:line])

    # Allow 'border' as a synonym for 'line'.
    line = line_properties(trendline[:border]) if trendline[:border]

    # Set the fill properties for the trendline.
    fill = fill_properties(trendline[:fill])

    trendline[:_line] = line
    trendline[:_fill] = fill

    return trendline
  end

  #
  # Convert user defined error bars properties to structure required
  # internally.
  #
  def error_bars_properties(params = {})
    return if !ptrue?(params) || params.empty?

    # Default values.
    error_bars = {
      :_type      => 'fixedVal',
      :_value     => 1,
      :_endcap    => 1,
      :_direction => 'both'
    }

    types = {
      :fixed              => 'fixedVal',
      :percentage         => 'percentage',
      :standard_deviation => 'stdDev',
      :standard_error     => 'stdErr'
    }

    # Check the error bars type.
    error_type = params[:type].to_sym

    if types.key?(error_type)
      error_bars[:_type] = types[error_type]
    else
      raise "Unknown error bars type '#{error_type}'\n"
    end

    # Set the value for error types that require it.
    if params.key?(:value)
      error_bars[:_value] = params[:value]
    end

    # Set the end-cap style.
    if params.key?(:end_style)
      error_bars[:_endcap] = params[:end_style]
    end

    # Set the error bar direction.
    if params.key?(:direction)
      if params[:direction] == 'minus'
        error_bars[:_direction] = 'minus'
      elsif params[:direction] == 'plus'
        error_bars[:_direction] = 'plus'
      else
        # Default to 'both'
      end
    end

    # Set the line properties for the error bars.
    error_bars[:_line] = line_properties(params[:line])

    error_bars
  end

  #
  # Convert user defined points properties to structure required internally.
  #
  def points_properties(user_points = nil)
    return unless user_points

    points = []
    user_points.each do |user_point|
      if user_point
        # Set the lline properties for the point.
        line = line_properties(user_point[:line])

        # Allow 'border' as a synonym for 'line'.
        if user_point[:border]
          line = line_properties(user_point[:border])
        end

        # Set the fill properties for the chartarea.
        fill = fill_properties(user_point[:fill])

        point = {}
        point[:_line] = line
        point[:_fill] = fill
      end
      points << point
    end
    points
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

  def value_or_raise(hash, key, msg)
    raise "Unknown #{msg} '#{key}'" unless hash[key.to_sym]
    hash[key.to_sym]
  end
end
