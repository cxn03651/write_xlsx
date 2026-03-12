# -*- coding: utf-8 -*-
# frozen_string_literal: true

require 'write_xlsx/chart/axis'
require 'write_xlsx/chart/legend'

module Writexlsx
  class Chart
    module Initialization
      private

      def axis_setup
        @axis_ids          = []
        @axis2_ids         = []
        @cat_has_num_fmt   = false
        @requires_category = 0
        @cat_axis_position = 'b'
        @val_axis_position = 'l'
        @horiz_cat_axis    = 0
        @horiz_val_axis    = 1
        @x_axis            = Axis.new(self)
        @y_axis            = Axis.new(self)
        @x2_axis           = Axis.new(self)
        @y2_axis           = Axis.new(self)
      end

      def display_setup
        @orientation       = 0x0
        @width             = 480
        @height            = 288
        @x_scale           = 1
        @y_scale           = 1
        @x_offset          = 0
        @y_offset          = 0
        @legend            = Legend.new
        @smooth_allowed    = 0
        @cross_between     = 'between'
        @date_category     = false
        @show_blanks       = 'gap'
        @show_na_as_empty  = false
        @show_hidden_data  = false
        @show_crosses      = true
      end

      #
      # Setup the default properties for a chart.
      #
      def set_default_properties # :nodoc:
        display_setup
        axis_setup
        set_axis_defaults

        set_x_axis
        set_y_axis

        set_x2_axis
        set_y2_axis
      end

      def set_axis_defaults
        @x_axis.defaults  = x_axis_defaults
        @y_axis.defaults  = y_axis_defaults
        @x2_axis.defaults = x2_axis_defaults
        @y2_axis.defaults = y2_axis_defaults
      end

      def x_axis_defaults
        {
          num_format:      'General',
          major_gridlines: { visible: 0 }
        }
      end

      def y_axis_defaults
        {
          num_format:      'General',
          major_gridlines: { visible: 1 }
        }
      end

      def x2_axis_defaults
        {
          num_format:     'General',
          label_position: 'none',
          crossing:       'max',
          visible:        0
        }
      end

      def y2_axis_defaults
        {
          num_format:      'General',
          major_gridlines: { visible: 0 },
          position:        'right',
          visible:         1
        }
      end
    end
  end
end
