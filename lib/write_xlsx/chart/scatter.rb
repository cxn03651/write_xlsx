# -*- coding: utf-8 -*-
###############################################################################
#
# Scatter - A class for writing Excel Scatter charts.
#
# Used in conjunction with Chart.
#
# See formatting note in Chart.
#
# Copyright 2000-2011, John McNamara, jmcnamara@cpan.org
# Convert to ruby by Hideo NAKAMURA, cxn03651@msj.biglobe.ne.jp
#

require 'write_xlsx/package/xml_writer_simple'
require 'write_xlsx/chart'
require 'write_xlsx/utility'

module Writexlsx
  class Chart
    # The Scatter chart module also supports the following sub-types:
    #
    #     markers_only (the default)
    #     straight_with_markers
    #     straight
    #     smooth_with_markers
    #     smooth
    # These can be specified at creation time via the add_chart() Worksheet
    # method:
    #
    #     chart = workbook.add_chart(
    #         :type    => 'scatter',
    #         :subtype => 'straight_with_markers'
    #     )
    #
    class Scatter < self
      include Writexlsx::Utility

      def initialize(subtype)
        super(subtype)
        @subtype           = subtype || 'marker_only'
        @cross_between     = 'midCat'
        @horiz_val_axis    = 0
        @val_axis_position = 'b'
        @smooth_allowed    = 1
      end

      #
      # Override the virtual superclass method with a chart specific method.
      #
      def write_chart_type(params)
        # Write the c:areaChart element.
        write_scatter_chart(params)
      end

      #
      # Write the <c:scatterChart> element.
      #
      def write_scatter_chart(params)
        if params[:primary_axes] == 1
          series = get_primary_axes_series
        else
          series = get_secondary_axes_series
        end
        return if series.empty?

        style   = 'lineMarker'
        subtype = @subtype

        # Set the user defined chart subtype
        case subtype
        when 'marker_only', 'straight_with_markers', 'straight'
          style = 'lineMarker'
        when 'smooth_with_markers', 'smooth'
          style = 'smoothMarker'
        end

        # Add default formatting to the series data.
        modify_series_formatting

        @writer.tag_elements('c:scatterChart') do
          # Write the c:scatterStyle element.
          write_scatter_style(style)
          # Write the series elements.
          series.each {|s| write_series(s)}

          # Write the c:marker element.
          write_marker_value

          # Write the c:axId elements
          write_axis_ids(params)
        end
      end

      #
      # Over-ridden to write c:xVal/c:yVal instead of c:cat/c:val elements.
      #
      # Write the <c:ser> element.
      #
      def write_ser(series)
        index = @series_index
        @series_index += 1

        @writer.tag_elements('c:ser') do
          # Write the c:idx element.
          write_idx(index)
          # Write the c:order element.
          write_order(index)
          # Write the series name.
          write_series_name(series)
          # Write the c:spPr element.
          write_sp_pr(series)
          # Write the c:marker element.
          write_marker(series.marker)
          # Write the c:dPt element.
          write_d_pt(series.points)
          # Write the c:dLbls element.
          write_d_lbls(series.labels)
          # Write the c:trendline element.
          write_trendline(series.trendline)
          # Write the c:errBars element.
          write_error_bars(series.error_bars)
          # Write the c:xVal element.
          write_x_val(series)
          # Write the c:yVal element.
          write_y_val(series)
          # Write the c:smooth element.
          if @subtype =~ /smooth/ && !series.smooth
            write_c_smooth(1)
          else
            write_c_smooth(series.smooth)
          end
        end
      end

      #
      # Over-ridden to have 2 valAx elements for scatter charts instead of
      # catAx/valAx.
      #
      # Write the <c:plotArea> element.
      #
      def write_plot_area
        @writer.tag_elements('c:plotArea') do
          # Write the c:layout element.
          write_layout(@plotarea.layout, 'plot')

          # Write the subclass chart type elements for primary and secondary axes
          write_chart_type(:primary_axes => 1)
          write_chart_type(:primary_axes => 0)

          # Write c:catAx and c:valAx elements for series using primary axes
          write_cat_val_axis(
                             :x_axis   => @x_axis,
                             :y_axis   => @y_axis,
                             :axis_ids => @axis_ids,
                             :position => 'b'
                             )
          tmp = @horiz_val_axis
          @horiz_val_axis = 1
          write_val_axis(
                         :x_axis   => @x_axis,
                         :y_axis   => @y_axis,
                         :axis_ids => @axis_ids,
                         :position => 'l'
                         )
          @horiz_val_axis = tmp

          # Write c:valAx and c:catAx elements for series using secondary axes
          write_cat_val_axis(
                             :x_axis   => @x2_axis,
                             :y_axis   => @y2_axis,
                             :axis_ids => @axis2_ids,
                             :position => 'b'
                             )
          @horiz_val_axis = 1
          write_val_axis(
                         :x_axis   => @x2_axis,
                         :y_axis   => @y2_axis,
                         :axis_ids => @axis2_ids,
                         :position => 'l'
                         )

          # Write the c:spPr element for the plotarea formatting.
          write_sp_pr(@plotarea)
        end
      end

      #
      # Write the <c:xVal> element.
      #
      def write_x_val(series)
        formula = series.categories
        data_id = series.cat_data_id
        data    = @formula_data[data_id]

        @writer.tag_elements('c:xVal') do
          # Check the type of cached data.
          type = get_data_type(data)

          # TODO. Can a scatter plot have non-numeric data.

          if type == 'str'
            # Write the c:numRef element.
            write_str_ref(formula, data, type)
          else
            write_num_ref(formula, data, type)
          end
        end
      end

      #
      # Write the <c:yVal> element.
      #
      def write_y_val(series)
        formula = series.values
        data_id = series.val_data_id
        data    = @formula_data[data_id]

        @writer.tag_elements('c:yVal') do
          # Unlike Cat axes data should only be numeric

          # Write the c:numRef element.
          write_num_ref(formula, data, 'num')
        end
      end

      #
      # Write the <c:scatterStyle> element.
      #
      def write_scatter_style(val)
        attributes = [ ['val', val] ]

        @writer.empty_tag('c:scatterStyle', attributes)
      end

      #
      # Add default formatting to the series data unless it has already been
      # specified by the user.
      #
      def modify_series_formatting
        subtype = @subtype

        # The default scatter style "markers only" requires a line type
        if subtype == 'marker_only'
          # Go through each series and define default values.
          @series.each do |series|
            # Set a line type unless there is already a user defined type.
            unless series.line_defined?
              series.line = line_properties(:width => 2.25, :none => 1, :_defined => 1)
            end
          end
        end

        # Turn markers off for subtypes that don't have them
        unless subtype =~ /marker/
          # Go through each series and define default values.
          @series.each do |series|
            # Set a marker type unless there is already a user defined type.
            unless ptrue?(series.marker)
              series.marker = Marker.new(:type => 'none', :_defined => 1)
            end
          end
        end
      end

      #
      # Write an individual <c:dPt> element. Override the parent method to add
      # markers.
      #
      def write_d_pt_point(index, point)
        @writer.tag_elements('c:dPt') do
          # Write the c:idx element.
          write_idx(index)
          @writer.tag_elements('c:marker') do
            # Write the c:spPr element.
            write_sp_pr(point)
          end
        end
      end
    end
  end
end
