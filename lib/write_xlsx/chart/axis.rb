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
      def convert_axis_args(chart, params) # :nodoc:
        @chart = chart

        axis = convert_axis_args2(self, params)

        [
         :_defaults, :_name, :_formula, :_data_id, :_reverse,
         :_min, :_max,
         :_minor_unit, :_major_unit, :_minor_unit_type, :_major_unit_type,
         :_log_base, :_crossing, :_position, :_label_position, :_visible,
         :_num_format, :_num_format_linked, :_num_font, :_name_font,
         :_major_gridlines, :_minor_gridlines, :_major_tick_mark
        ].each do |m|
          instance_variable_set(
                                "@#{m}",
                                axis.instance_variable_get("@#{m}")
                                )
        end
=begin
           @chart = chart
        args = (_defaults || {}).merge(params)

        @_name, @_name_formula =
          @chart.process_names(args[:name], args[:name_formula])
        @_data_id = @chart.get_data_id(@_name_formula, args[:data])
        @_reverse = args[:reverse]
        @_min = args[:min]
        @_max = args[:max]
        @_minor_unit = args[:minor_unit]
        @_major_unit = args[:major_unit]
        @_minor_unit_type = args[:minor_unit_type]
        @_major_unit_type = args[:major_unit_type]
        @_log_base = args[:log_base]
        @_crossing = args[:crossing]
        @_label_position = args[:label_position]
        @_num_format = args[:num_format]
        @_num_format_linked = args[:num_format_linked]
        @_visible           = args[:visible] || 1

        # Only use the first letter of bottom, top, left or right.
        @_position = args[:position]
        @_position = @_position.downcase[0, 1] if @_position

        # Map major/minor_gridlines properties.
        if args[:major_gridlines] && ptrue?(args[:major_gridlines][:visible])
          @_major_gridlines = get_gridline_properties(args[:major_gridlines])
        else
          @_major_gridlines = nil
        end
        if args[:minor_gridlines] && ptrue?(args[:minor_gridlines][:visible])
          @_minor_gridlines = get_gridline_properties(args[:minor_gridlines])
        end

        # Set the font properties if present.
        @_num_font  = @chart.convert_font_args(args[:num_font])
        @_name_font = @chart.convert_font_args(args[:name_font])
=end
      end

      #
      # Convert user defined axis values into private hash values.
      #
      def convert_axis_args2(axis, params) # :nodoc:
        arg = (axis._defaults || {}).merge(params)
        name, name_formula = @chart.process_names(arg[:name], arg[:name_formula])

        data_id = @chart.get_data_id(name_formula, arg[:data])

        a = Axis.new
        a._defaults          = axis._defaults
        a._name              = name
        a._formula           = name_formula
        a._data_id           = data_id
        a._reverse           = arg[:reverse]
        a._min               = arg[:min]
        a._max               = arg[:max]
        a._minor_unit        = arg[:minor_unit]
        a._major_unit        = arg[:major_unit]
        a._minor_unit_type   = arg[:minor_unit_type]
        a._major_unit_type   = arg[:major_unit_type]
        a._log_base          = arg[:log_base]
        a._crossing          = arg[:crossing]
        a._position          = arg[:position]
        a._label_position    = arg[:label_position]
        a._num_format        = arg[:num_format]
        a._num_format_linked = arg[:num_format_linked]
        a._visible           = arg[:visible] || 1

        # Map major/minor_gridlines properties.
        [:major_gridlines, :minor_gridlines].each do |lines|
          if arg[lines] && ptrue?(arg[lines][:visible])
            a.instance_variable_set("@_#{lines}", get_gridline_properties(arg[lines]))
          end
        end

        # Only use the first letter of bottom, top, left or right.
        a._position = a._position.downcase[0, 1] if a._position

        # Set the font properties if present.
        a._num_font = @chart.convert_font_args(arg[:num_font])
        a._name_font  = @chart.convert_font_args(arg[:name_font])

        a
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
