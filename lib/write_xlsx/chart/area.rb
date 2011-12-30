# -*- coding: utf-8 -*-
###############################################################################
#
# Area - A class for writing Excel Area charts.
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
    class Area < self
      include Writexlsx::Utility

      def initialize
        super(self.class)
        @cross_between = 'midCat'
      end

      #
      # Override the virtual superclass method with a chart specific method.
      #
      def write_chart_type
        # Write the c:areaChart element.
        write_area_chart
      end

      #
      # Write the <c:areaChart> element.
      #
      def write_area_chart
        @writer.start_tag('c:areaChart')

        # Write the c:grouping element.
        write_grouping('standard')

        # Write the series elements.
        write_series

        @writer.end_tag('c:areaChart')
      end
    end
  end
end
