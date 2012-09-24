# -*- coding: utf-8 -*-
###############################################################################
#
# Line - A class for writing Excel Line charts.
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
    class Line < self
      include Writexlsx::Utility

      def initialize(subtype)
        super(subtype)
        @default_marker = {:type => 'none'}
      end

      #
      # Override the virtual superclass method with a chart specific method.
      #
      def write_chart_type(params)
        # Write the c:barChart element.
        write_line_chart(params)
      end

      #
      # Write the <c:lineChart> element.
      #
      def write_line_chart(params)
        if params[:primary_axes] != 0
          series = get_primary_axes_series
        else
          series = get_secondary_axes_series
        end

        return if series.empty?

        @writer.tag_elements('c:lineChart') do
          # Write the c:grouping element.
          write_grouping('standard')
          # Write the series elements.
          series.each {|s| write_series(s)}

          # Write the c:marker element.
          write_marker_value

          # Write the c:axId elements
          write_axis_ids(params)
        end
      end
    end
  end
end
