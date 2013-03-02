# -*- encoding: utf-8 -*-
require 'write_xlsx/utility'

module Writexlsx
  class Chart
    class Axis
      include Writexlsx::Utility

      attr_accessor :_defaults, :_name, :_formula, :_data_id, :_reverse
      attr_accessor  :_min, :_max
      attr_accessor :_minor_unit, :_major_unit, :_minor_unit_type, :_major_unit_type
      attr_accessor :_log_base, :_crossing, :_position, :_label_position, :_visible
      attr_accessor :_num_format, :_num_format_linked, :_num_font, :_name_font
      attr_accessor :_major_gridlines, :_minor_gridlines, :_major_tick_mark

      #
      # Convert user defined axis values into axis instance.
      #
      def merge_with_hash(chart, params) # :nodoc:
        @chart = chart
        @args = args = (_defaults || {}).merge(params)
        @_name, @_formula = @chart.process_names(@args[:name], @args[:name_formula])
        @_data_id           = @chart.get_data_id(@_formula, @args[:data])
        @_reverse           = @args[:reverse]
        @_min               = @args[:min]
        @_max               = @args[:max]
        @_minor_unit        = @args[:minor_unit]
        @_major_unit        = @args[:major_unit]
        @_minor_unit_type   = @args[:minor_unit_type]
        @_major_unit_type   = @args[:major_unit_type]
        @_log_base          = @args[:log_base]
        @_crossing          = @args[:crossing]
        @_label_position    = @args[:label_position]
        @_num_format        = @args[:num_format]
        @_num_format_linked = @args[:num_format_linked]
        @_visible           = @args[:visible] || 1

        # Map major/minor_gridlines properties.
        [:major_gridlines, :minor_gridlines].each do |lines|
          if @args[lines] && ptrue?(@args[lines][:visible])
            instance_variable_set("@_#{lines}", get_gridline_properties(@args[lines]))
          else
            instance_variable_set("@_#{lines}", nil)
          end
        end
        @_major_tick_mark   = @args[:major_tick_mark]

        # Only use the first letter of bottom, top, left or right.
        @_position = @args[:position]
        @_position = @_position.downcase[0, 1] if @_position

        # Set the font properties if present.
        @_num_font  = @chart.convert_font_args(@args[:num_font])
        @_name_font = @chart.convert_font_args(@args[:name_font])
      end

      def set_property(params, property)
        instance_variable_set(
                              "@_#{property}",
                              params[property] || @_defaults[property]
                              )
      end

      #
      # Convert user defined gridline properties to the structure required internally.
      #
      def get_gridline_properties(args)
        # Set the visible property for the gridline.
        gridline = { :_visible => args[:visible] }

        # Set the line properties for the gridline.
        gridline[:_line] = @chart.get_line_properties(args[:line])

        gridline
      end
    end
  end
end
