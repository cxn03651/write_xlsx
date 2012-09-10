# -*- coding: utf-8 -*-
###############################################################################
#
# Column - A class for writing Excel Column charts.
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
    # The Column chart module also supports the following sub-types:
    #
    #     stacked
    #     percent_stacked
    # These can be specified at creation time via the add_chart() Worksheet
    # method:
    #
    #     chart = workbook.add_chart( :type => 'column', :subtype => 'stacked' )
    #
    class Column < self
      include Writexlsx::Utility

      def initialize(subtype)
        super(subtype)
        @subtype = 'clustered'
        @horiz_val_axis = 0
      end

      #
      # Override the virtual superclass method with a chart specific method.
      #
      def write_chart_type
        # Write the c:barChart element.
        write_bar_chart
      end

      #
      # Write the <c:barDir> element.
      #
      def write_bar_dir
        val  = 'col'

        attributes = ['val', val]

        @writer.empty_tag('c:barDir', attributes)
      end


      #
      # Over-ridden to add c:overlap.
      #
      # Write the series elements.
      #
      def write_series
        write_series_base {write_overlap if @subtype =~ /stacked/}
      end

      #
      # Over-ridden to add % format. TODO. This will be refactored back up to the
      # SUPER class later.
      #
      # Write the <c:numFmt> element.
      #
      def write_number_format(format_code = nil)
        format_code   ||= 'General'
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
