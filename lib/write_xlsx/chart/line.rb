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

      def initialize
        super(self.class)
        @default_marker = {:type => 'none'}
      end

      #
      # Override the virtual superclass method with a chart specific method.
      #
      def write_chart_type
        # Write the c:barChart element.
        write_line_chart
      end

      #
      # Write the <c:lineChart> element.
      #
      def write_line_chart
        @writer.start_tag('c:lineChart')

        # Write the c:grouping element.
        write_grouping('standard')

        # Write the series elements.
        write_series

        @writer.end_tag('c:lineChart')
      end
    end
  end
end
