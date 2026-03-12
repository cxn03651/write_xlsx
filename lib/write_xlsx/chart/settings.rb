# -*- coding: utf-8 -*-
# frozen_string_literal: true

###############################################################################
#
# settings.rb - common chart facade and orchestration
#
###############################################################################

module Writexlsx
  class Chart
    module Settings
      #
      # Add a series and it's properties to a chart.
      #
      def add_series(params)
        # Check that the required input has been specified.
        raise "Must specify ':values' in add_series" unless params.has_key?(:values)

        raise "Must specify ':categories' in add_series for this chart type" if @requires_category != 0 && !params.has_key?(:categories)

        raise "The maximum number of series that can be added to an Excel Chart is 255." if @series.size == 255

        @series << Series.new(self, params)

        # Set the secondary axis properties.
        x2_axis = params[:x2_axis]
        y2_axis = params[:y2_axis]

        # Store secondary status for combined charts.
        @is_secondary = true if ptrue?(x2_axis) || ptrue?(y2_axis)

        # Set the gap and overlap for Bar/Column charts.
        if params[:gap]
          if ptrue?(y2_axis)
            @series_gap_2 = params[:gap]
          else
            @series_gap_1 = params[:gap]
          end
        end

        # Set the overlap for Bar/Column charts.
        if params[:overlap]
          if ptrue?(y2_axis)
            @series_overlap_2 = params[:overlap]
          else
            @series_overlap_1 = params[:overlap]
          end
        end
      end

      #
      # Set the properties of the x-axis.
      #
      def set_x_axis(params = {})
        @date_category = true if ptrue?(params[:date_axis])
        @x_axis.merge_with_hash(params)
      end

      #
      # Set the properties of the Y-axis.
      #
      # The set_y_axis() method is used to set properties of the Y axis.
      # The properties that can be set are the same as for set_x_axis,
      #
      def set_y_axis(params = {})
        @date_category = true if ptrue?(params[:date_axis])
        @y_axis.merge_with_hash(params)
      end

      #
      # Set the properties of the secondary X-axis.
      #
      def set_x2_axis(params = {})
        @date_category = true if ptrue?(params[:date_axis])
        @x2_axis.merge_with_hash(params)
      end

      #
      # Set the properties of the secondary Y-axis.
      #
      def set_y2_axis(params = {})
        @date_category = true if ptrue?(params[:date_axis])
        @y2_axis.merge_with_hash(params)
      end

      #
      # Set the properties of the chart title.
      #
      def set_title(params)
        @title.merge_with_hash(params)
      end

      #
      # Set the properties of the chart legend.
      #
      def set_legend(params)
        # Convert the user default properties to internal properties.
        legend_properties(params)
      end

      #
      # Set the properties of the chart plotarea.
      #
      def set_plotarea(params)
        # Convert the user defined properties to internal properties.
        @plotarea = ChartArea.new(params)
      end

      #
      # Set the properties of the chart chartarea.
      #
      def set_chartarea(params)
        # Convert the user defined properties to internal properties.
        @chartarea = ChartArea.new(params)
      end

      #
      # Set on of the 42 built-in Excel chart styles. The default style is 2.
      #
      def set_style(style_id = 2)
        style_id = 2 if style_id < 1 || style_id > 48
        @style_id = style_id
      end

      #
      # Set the option for displaying blank data in a chart. The default is 'gap'.
      #
      def show_blanks_as(option)
        return unless option

        raise "Unknown show_blanks_as() option '#{option}'\n" unless %i[gap zero span].include?(option.to_sym)

        @show_blanks = option
      end

      #
      # Display data in hidden rows or columns on the chart.
      #
      def show_hidden_data
        @show_hidden_data = true
      end

      #
      # Set dimensions for scale for the chart.
      #
      def set_size(params = {})
        @width    = params[:width]    if params[:width]
        @height   = params[:height]   if params[:height]
        @x_scale  = params[:x_scale]  if params[:x_scale]
        @y_scale  = params[:y_scale]  if params[:y_scale]
        @x_offset = params[:x_offset] if params[:x_offset]
        @y_offset = params[:y_offset] if params[:y_offset]
      end

      # Backward compatibility with poorly chosen method name.
      alias size set_size

      #
      # The set_table method adds a data table below the horizontal axis with the
      # data used to plot the chart.
      #
      def set_table(params = {})
        @table = Table.new(params)
        @table.palette = @palette
      end

      #
      # Set properties for the chart up-down bars.
      #
      def set_up_down_bars(params = {})
        # Map border to line.
        %i[up down].each do |up_down|
          if params[up_down]
            params[up_down][:line] = params[up_down][:border] if params[up_down][:border]
          else
            params[up_down] = {}
          end
        end

        # Set the up and down bar properties.
        @up_down_bars = {
          _up:   Chartline.new(params[:up]),
          _down: Chartline.new(params[:down])
        }
      end

      #
      # Set properties for the chart drop lines.
      #
      def set_drop_lines(params = {})
        @drop_lines = Chartline.new(params)
      end

      #
      # Set properties for the chart high-low lines.
      #
      def set_high_low_lines(params = {})
        @hi_low_lines = Chartline.new(params)
      end

      #
      # Add another chart to create a combined chart.
      #
      def combine(chart)
        @combined = chart
      end

      #
      # Setup the default configuration data for an embedded chart.
      #
      def set_embedded_config_data
        @embedded = true
      end

      #
      # Set the option for displaying #N/A as an empty cell in a chart.
      #
      def show_na_as_empty_cell
        @show_na_as_empty = true
      end
    end
  end
end
