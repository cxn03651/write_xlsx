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

      def initialize(subtype)
        super(subtype)
        @subtype = subtype || 'standard'
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
        if @subtype == 'percent_stacked'
          subtype = 'percentStacked'
        else
          subtype = @subtype
        end
        @writer.tag_elements('c:areaChart') do
          # Write the c:grouping element.
          write_grouping(subtype)
          # Write the series elements.
          write_series
        end
      end

      #
      # Over-ridden to add % format. TODO. This will be refactored back up to the
      # SUPER class later.
      #
      # Write the <C:numFmt> element.
      #
      def write_number_format(format_code = nil)
        source_linked = 1
        format_code = 'General' if !format_code || format_code.empty?
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
