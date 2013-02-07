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
      include Writexlsx::Utility

      def initialize(subtype)
        super(subtype)
        @subtype = subtype || 'clustered'
        @cat_axis_position = 'l'
        @val_axis_position = 'b'
        @horiz_val_axis    = 0
        @horiz_cat_axis    = 1
        @show_crosses      = false

        # Override and reset the default axis values.
        if @x_axis[:_defaults]
          @x_axis[:_defaults][:major_gridlines] = { :visible => 1 }
        else
          @x_axis[:_defaults] = { :major_gridlines => { :visible => 1 } }
        end
        if @y_axis[:_defaults]
          @y_axis[:_defaults][:major_gridlines] = { :visible => 0 }
        else
          @y_axis[:_defaults] = { :major_gridlines => { :visible => 0 } }
        end

        if @subtype == 'percent_stacked'
            @x_axis[:_defaults][:num_format] = '0%'
        end

        set_x_axis
        set_y_axis
      end

      #
      # Override the virtual superclass method with a chart specific method.
      #
      def write_chart_type(params)
        if params[:primary_axes] != 0
          # Reverse X and Y axes for Bar charts.
          @y_axis, @x_axis = @x_axis, @y_axis
          if @y2_axis[:_position] == 'r'
            @y2_axis[:_position] = 't'
          end
        end

        # Write the c:barChart element.
        write_bar_chart(params)
      end

      #
      # Write the <c:barDir> element.
      #
      def write_bar_dir
        val  = 'bar'

        attributes = ['val', val]

        @writer.empty_tag('c:barDir', attributes)
      end
    end
  end
end
