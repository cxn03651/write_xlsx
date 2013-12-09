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
    # A Pie chart doesn't have an X or Y axis so the following common chart
    # methods are ignored.
    #
    #     chart.set_x_axis
    #     chart.set_y_axis
    #
    class Pie < self
      include Writexlsx::Utility

      def initialize(subtype)
        super(subtype)
        @vary_data_color = 1
      end

      #
      # Override the virtual superclass method with a chart specific method.
      #
      def write_chart_type(params = {})
        # Write the c:areaChart element.
        write_pie_chart
      end

      #
      # Write the <c:pieChart> element. Over-ridden method to remove axis_id code
      # since pie charts don't require val and vat axes.
      #
      def write_pie_chart
        @writer.tag_elements('c:pieChart') do
          # Write the c:varyColors element.
          write_vary_colors
          # Write the series elements.
          @series.each {|s| write_series(s)}
          # Write the c:firstSliceAng element.
          write_first_slice_ang
        end
      end

      #
      # Over-ridden method to remove the cat_axis() and val_axis() code since
      # Pie charts don't require those axes.
      #
      # Write the <c:plotArea> element.
      #
      def write_plot_area
        @writer.tag_elements('c:plotArea') do
          # Write the c:layout element.
          write_layout(@plotarea[:_layout], 'plot')
          # Write the subclass chart type element.
          write_chart_type
        end
      end

      #
      # Over-ridden method to add <c:txPr> to legend.
      #
      # Write the <c:legend> element.
      #
      def write_legend
        position = @legend_position
        overlay  = 0

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

        @writer.tag_elements('c:legend') do
          # Write the c:legendPos element.
          write_legend_pos(position)
          # Write the c:layout element.
          write_layout(@legend_layout, 'legend')
          # Write the c:overlay element.
          write_overlay if overlay != 0
          # Write the c:txPr element. Over-ridden.
          write_tx_pr_legend
        end
      end

      #
      # Write the <c:txPr> element for legends.
      #
      def write_tx_pr_legend
        horiz = 0

        @writer.tag_elements('c:txPr') do
          # Write the a:bodyPr element.
          write_a_body_pr(nil, horiz)
          # Write the a:lstStyle element.
          write_a_lst_style
          # Write the a:p element.
          write_a_p_legend
        end
      end

      #
      # Write the <a:p> element for legends.
      #
      def write_a_p_legend
        @writer.tag_elements('a:p') do
          # Write the a:pPr element.
          write_a_p_pr_legend
          # Write the a:endParaRPr element.
          write_a_end_para_rpr
        end
      end

      #
      # Write the <a:pPr> element for legends.
      #
      def write_a_p_pr_legend
        rtl  = 0

        attributes = [ ['rtl', rtl] ]

        @writer.tag_elements('a:pPr', attributes) do
          # Write the a:defRPr element.
          write_a_def_rpr
        end
      end

      #
      # Write the <c:varyColors> element.
      #
      def write_vary_colors
        val  = 1

        attributes = [ ['val', val] ]

        @writer.empty_tag('c:varyColors', attributes)
      end

      #
      # Write the <c:firstSliceAng> element.
      #
      def write_first_slice_ang
        val  = 0

        attributes = [ ['val', val] ]

        @writer.empty_tag('c:firstSliceAng', attributes)
      end
    end
  end
end
