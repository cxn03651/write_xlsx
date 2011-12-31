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

      def initialize
        super(self.class)
      end

      #
      # Override the virtual superclass method with a chart specific method.
      #
      def write_chart_type
        # Write the c:areaChart element.
        write_stock_chart
      end

      #
      # Write the <c:stockChart> element.
      #
      def write_stock_chart
        # Add default formatting to the series data.
        modify_series_formatting

        @writer.start_tag('c:stockChart')

        # Write the series elements.
        write_series

        @writer.end_tag('c:stockChart')
      end

      #
      # Over-ridden to add hi_low_lines(). TODO. Refactor up into the SUPER class.
      #
      # Write the series elements.
      #
      def write_series
        # Write each series with subelements.
        index = 0
        @series.each do |series|
          write_ser(index, series)
          index += 1
        end

        # Write the c:hiLowLines element.
        write_hi_low_lines

        # Write the c:marker element.
        write_marker_value

        # Generate the axis ids.
        add_axis_id
        add_axis_id

        # Write the c:axId element.
        write_axis_id(@axis_ids[0])
        write_axis_id(@axis_ids[1])
      end

      #
      # Write the <c:plotArea> element.
      #
      def write_plot_area
        @writer.start_tag('c:plotArea')

        # Write the c:layout element.
        write_layout

        # Write the subclass chart type element.
        write_chart_type

        # Write the c:dateAx element.
        write_date_axis

        # Write the c:catAx element.
        write_val_axis

        @writer.end_tag('c:plotArea')
      end

      #
      # Add default formatting to the series data.
      #
      def modify_series_formatting
        index = 0
        array = []
        @series.each do |series|
          if index % 4 != 3
            if series[:_line][:_defined].nil? || series[:_line][:_defined] == 0
              series[:_line] = {
                :width    => 2.25,
                :none     => 1,
                :_defined => 1
              }
            end

            if series[:_marker].nil? || series[:_marker] == 0
              if index % 4 == 2
                series[:_marker] = { :type => 'dot', :size => 3 }
              else
                series[:_marker] = { :type => 'none' }
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
