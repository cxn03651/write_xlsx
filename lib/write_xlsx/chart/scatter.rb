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
    class Scatter < self
      include Utility

      def initialize
        super(self.class)
        @cross_between = 'midCat'
        @horiz_val_axis    = 0
      end

      #
      # Override the virtual superclass method with a chart specific method.
      #
      def write_chart_type
        # Write the c:areaChart element.
        write_scatter_chart
      end

      #
      # Write the <c:pieChart> element.
      #
      def write_scatter_chart
        # Add default formatting to the series data.
        modify_series_formatting

        @writer.start_tag('c:scatterChart')

        # Write the c:scatterStyle element.
        write_scatter_style

        # Write the series elements.
        write_series

        @writer.end_tag('c:scatterChart')
      end

      #
      # Over-ridden to write c:xVal/c:yVal instead of c:cat/c:val elements.
      #
      # Write the <c:ser> element.
      #
      def write_ser(index, series)
        @writer.start_tag('c:ser')

        # Write the c:idx element.
        write_idx(index)

        # Write the c:order element.
        write_order(index)

        # Write the series name.
        write_series_name(series)

        # Write the c:spPr element.
        write_sp_pr(series)

        # Write the c:marker element.
        write_marker(series[:_marker])

        # Write the c:dLbls element.
        write_d_lbls(series[:_labels])

        # Write the c:trendline element.
        write_trendline(series[:_trendline])

        # Write the c:xVal element.
        write_x_val(series)

        # Write the c:yVal element.
        write_y_val(series)

        @writer.end_tag('c:ser')
      end

      #
      # Over-ridden to have 2 valAx elements for scatter charts instead of
      # catAx/valAx.
      #
      # Write the <c:plotArea> element.
      #
      def write_plot_area
        @writer.start_tag('c:plotArea')

        # Write the c:layout element.
        write_layout

        # Write the subclass chart type element.
        write_chart_type

        # Write the c:catAx element.
        write_cat_val_axis('b', 1)

        # Write the c:catAx element.
        @horiz_val_axis = 1
        write_val_axis('l')

        @writer.end_tag('c:plotArea')
      end

      #
      # Write the <c:xVal> element.
      #
      def write_x_val(series)
        formula = series[:_categories]
        data_id = series[:_cat_data_id]
        data    = @formula_data[data_id]

        @writer.start_tag('c:xVal')

        # Check the type of cached data.
        type = get_data_type(data)

        # TODO. Can a scatter plot have non-numeric data.

        if type == 'str'
          # Write the c:numRef element.
          write_str_ref(formula, data, type)
        else
          # Write the c:numRef element.
          write_num_ref(formula, data, type)
        end

        @writer.end_tag('c:xVal')
      end

      #
      # Write the <c:yVal> element.
      #
      def write_y_val(series)
        formula = series[:_values]
        data_id = series[:_val_data_id]
        data    = @formula_data[data_id]

        @writer.start_tag('c:yVal')

        # Check the type of cached data.
        type = get_data_type(data)

        if type == 'str'
          # Write the c:numRef element.
          write_str_ref(formula, data, type)
        else
          # Write the c:numRef element.
          write_num_ref(formula, data, type)
        end

        @writer.end_tag('c:yVal')
      end

      #
      # Write the <c:scatterStyle> element.
      #
      def write_scatter_style
        val  = 'lineMarker'

        attributes = ['val', val]

        @writer.empty_tag('c:scatterStyle', attributes)
      end

      #
      # Add default formatting to the series data.
      #
      def modify_series_formatting
        @series.collect! do |series|
          if series[:_line][:_defined].nil? || series[:_line][:_defined] == 0
              series[:_line] = {
                  :width    => 2.25,
                  :none     => 1,
                  :_defined => 1
              }
          end
          series
        end
      end
    end
  end
end
