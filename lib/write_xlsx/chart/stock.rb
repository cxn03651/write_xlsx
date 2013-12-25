# -*- coding: utf-8 -*-
###############################################################################
#
# Stock - A class for writing Excel Stock charts.
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
    #
    # The default Stock chart is an High-Low-Close chart.
    # A series must be added for each of these data sources.
    #
    class Stock < self
      include Writexlsx::Utility

      def initialize(subtype)
        super(subtype)
        @show_crosses = false
        @hi_low_lines = {}

        # Override and reset the default axis values.
        @x_axis.defaults[:num_format] = 'dd/mm/yyyy'
        @x2_axis.defaults[:num_format] = 'dd/mm/yyyy'
        set_x_axis
        set_x2_axis
      end

      #
      # Override the virtual superclass method with a chart specific method.
      #
      def write_chart_type(params)
        # Write the c:areaChart element.
        write_stock_chart(params)
      end

      #
      # Write the <c:stockChart> element.
      # Overridden to add hi_low_lines(). TODO. Refactor up into the SUPER class
      #
      def write_stock_chart(params)
        if params[:primary_axes] == 1
          series = get_primary_axes_series
        else
          series = get_secondary_axes_series
        end
        return if series.empty?

        # Add default formatting to the series data.
        modify_series_formatting

        @writer.tag_elements('c:stockChart') do
          # Write the series elements.
          series.each {|s| write_series(s)}

          # Write the c:dtopLines element.
          write_drop_lines

          # Write the c:hiLowLines element.
          write_hi_low_lines if ptrue?(params[:primary_axes])

          # Write the c:upDownBars element.
          write_up_down_bars

          # Write the c:marker element.
          write_marker_value

          # Write the c:axId elements
          write_axis_ids(params)
        end
      end

      #
      # Overridden to use write_date_axis() instead of write_cat_axis().
      #
      def write_plot_area
        write_plot_area_base(:stock)
      end

      #
      # Add default formatting to the series data.
      #
      def modify_series_formatting
        index = 0
        array = []
        @series.each do |series|
          if index % 4 != 3
            if series.line[:_defined].nil? || series.line[:_defined] == 0
              series.line = {
                :width    => 2.25,
                :none     => 1,
                :_defined => 1
              }
            end

            if series.marker.nil? || series.marker == 0
              if index % 4 == 2
                series.marker = Marker.new(:type => 'dot', :size => 3)
              else
                series.marker = Marker.new(:type => 'none')
              end
            end
          end
          index += 1
          array << series
        end
        @series = array
      end
    end
  end
end
