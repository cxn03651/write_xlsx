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
        series = axes_series(params)
        return if series.empty?

        # Add default formatting to the series data.
        modify_series_formatting

        @writer.tag_elements('c:stockChart') do
          # Write the series elements.
          @series.each {|s| write_series(s)}

          # Write the c:hiLowLines element.
          write_hi_low_lines if params[:primary_axes]

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
        @writer.tag_elements('c:plotArea') do
          # Write the c:layout element.
          write_layout

          # TODO: (for JMCNAMARA todo ;)
          # foreach my $chart_type (@chart_types)

          # Write the subclass chart type elements for primary and secondary axes
          write_chart_type(:primary_axes => 1)
          write_chart_type(:primary_axes => 0)

          # Write the c:dateAx elements for series using primary axes
          write_date_axis(
                          :x_axis   => @x_axis,
                          :y_axis   => @y_axis,
                          :axis_ids => @axis_ids
                          )
          write_val_axis(
                         :x_axis   => @x_axis,
                         :y_axis   => @y_axis,
                         :axis_ids => @axis_ids
                         )

          # Write c:valAx and c:catAx elements for series using secondary axes
          write_val_axis(
                         :x_axis   => @x2_axis,
                         :y_axis   => @y2_axis,
                         :axis_ids => @axis2_ids
                         )
          write_date_axis(
                          :x_axis   => @x2_axis,
                          :y_axis   => @y2_axis,
                          :axis_ids => @axis2_ids
                          )
        end
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
