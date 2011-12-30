# -*- coding: utf-8 -*-
###############################################################################
#
# Pie - A class for writing Excel Pie charts.
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
    class Pie < self
      include Writexlsx::Utility

      def initialize
        super(self.class)
        @vary_data_color = 1
      end

      #
      # Override the virtual superclass method with a chart specific method.
      #
      def write_chart_type
        # Write the c:areaChart element.
        write_pie_chart
      end

      #
      # Write the <c:pieChart> element.
      #
      def write_pie_chart
        @writer.start_tag('c:pieChart')

        # Write the c:varyColors element.
        write_vary_colors

        # Write the series elements.
        write_series

        # Write the c:firstSliceAng element.
        write_first_slice_ang

        @writer.end_tag('c:pieChart')
      end

      #
      # Over-ridden method to remove the cat_axis() and val_axis() code since
      # Pie charts don't require those axes.
      #
      # Write the <c:plotArea> element.
      #
      def write_plot_area
        @writer.start_tag('c:plotArea')

        # Write the c:layout element.
        write_layout

        # Write the subclass chart type element.
        write_chart_type

        @writer.end_tag('c:plotArea')
      end

      #
      # Over-ridden method to remove axis_id code since Pie charts  don't require
      # val and cat axes.
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
      end

      #
      # Over-ridden method to add <c:txPr> to legend.
      #
      # Write the <c:legend> element.
      #
      def write_legend
        position = @legend_position
        overlay = 0

        if position =~ /^overlay_/
          positon.sub!(/^overlay_/, '')
          overlay = 1
        end

        allowed = {
            'right'  => 'r',
            'left'   => 'l',
            'top'    => 't',
            'bottom' => 'b'
        }

        return if position == 'none'
        return unless allowed.has_key?(position)

        position = allowed[position]

        @writer.start_tag('c:legend')

        # Write the c:legendPos element.
        write_legend_pos(position)

        # Write the c:layout element.
        write_layout

        # Write the c:overlay element.
        write_overlay if overlay != 0

        # Write the c:txPr element. Over-ridden.
        write_tx_pr_legend

        @writer.end_tag('c:legend')
      end

      #
      # Write the <c:txPr> element for legends.
      #
      def write_tx_pr_legend
        horiz = 0

        @writer.start_tag('c:txPr')

        # Write the a:bodyPr element.
        write_a_body_pr(horiz)

        # Write the a:lstStyle element.
        write_a_lst_style

        # Write the a:p element.
        write_a_p_legend

        @writer.end_tag('c:txPr')
      end

      #
      # Write the <a:p> element for legends.
      #
      def write_a_p_legend
        @writer.start_tag('a:p')

        # Write the a:pPr element.
        write_a_p_pr_legend

        # Write the a:endParaRPr element.
        write_a_end_para_rpr

        @writer.end_tag('a:p')
      end

      #
      # Write the <a:pPr> element for legends.
      #
      def write_a_p_pr_legend
        rtl  = 0

        attributes = ['rtl', rtl]

        @writer.start_tag('a:pPr', attributes)

        # Write the a:defRPr element.
        write_a_def_rpr

        @writer.end_tag('a:pPr')
      end

      #
      # Write the <c:varyColors> element.
      #
      def write_vary_colors
        val  = 1

        attributes = ['val', val]

        @writer.empty_tag('c:varyColors', attributes)
      end

      #
      # Write the <c:firstSliceAng> element.
      #
      def write_first_slice_ang
        val  = 0

        attributes = ['val', val]

        @writer.empty_tag('c:firstSliceAng', attributes)
      end
    end
  end
end
