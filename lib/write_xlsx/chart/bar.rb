# -*- coding: utf-8 -*-
###############################################################################
#
# Bar - A class for writing Excel Bar charts.
#
# Used in conjunction with Chart.
#
# See formatting note in Chart.
#
# Copyright 2000-2011, John McNamara, jmcnamara@cpan.org
# Convert to ruby by Hideo NAKAMURA, cxn03651@msj.biglobe.ne.jp
#

require 'write_xlsx/package/xml_writer_simple'
require 'write_xlsx/utility'

module Writexlsx
  class Chart
    class Bar < self
      include Utility

      def initialize
        super(self.class)
        @subtype = 'clustered'
        @cat_axis_position = 'l'
        @val_axis_position = 'b'
        @horiz_val_axis    = 0
        @horiz_cat_axis    = 1
      end

      #
      # Override the virtual superclass method with a chart specific method.
      #
      def write_chart_type
        # Reverse X and Y axes for Bar charts.
        @x_axis, @y_axis = @y_axis, @x_axis

        # Write the c:barChart element.
        write_bar_chart
      end

      #
      # Write the <c:barChart> element.
      #
      def write_bar_chart
        subtype = @subtype

        subtype = 'percentStacked' if subtype == 'percent_stacked'

        @writer.start_tag('c:barChart')

        # Write the c:barDir element.
        write_bar_dir

        # Write the c:grouping element.
        write_grouping(subtype)

        # Write the series elements.
        write_series

        @writer.end_tag('c:barChart')
      end


      #
      # Write the <c:barDir> element.
      #
      def write_bar_dir
        val  = 'bar'

        attributes = ['val', val]

        @writer.empty_tag('c:barDir', attributes)
      end

      #
      # Over-ridden to add c:overlap.
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


        # Write the c:marker element.
        write_marker_value

        # Write the c:overlap element.
        write_overlap if @subtype =~ /stacked/

        # Generate the axis ids.
        add_axis_id
        add_axis_id

        # Write the c:axId element.
        write_axis_id(@axis_ids[0])
        write_axis_id(@axis_ids[1])
      end


      #
      # Over-ridden to add % format. TODO. This will be refactored back up to the
      # SUPER class later.
      #
      # Write the <c:numFmt> element.
      #
      def write_number_format(format_code = 'General')
        source_linked = 1

        format_code = '0%' if @subtype == 'percent_stacked'

        attributes = [
                      'formatCode',   format_code,
                      'sourceLinked', source_linked
                     ]

        @writer.empty_tag('c:numFmt', attributes)
      end
    end
  end
end
