# -*- coding: utf-8 -*-
# frozen_string_literal: true

module Writexlsx
  module Utility
    module ChartFormatting
      PATTERN_TYPES = {
        'percent_5'                => 'pct5',
        'percent_10'               => 'pct10',
        'percent_20'               => 'pct20',
        'percent_25'               => 'pct25',
        'percent_30'               => 'pct30',
        'percent_40'               => 'pct40',

        'percent_50'               => 'pct50',
        'percent_60'               => 'pct60',
        'percent_70'               => 'pct70',
        'percent_75'               => 'pct75',
        'percent_80'               => 'pct80',
        'percent_90'               => 'pct90',

        'light_downward_diagonal'  => 'ltDnDiag',
        'light_upward_diagonal'    => 'ltUpDiag',
        'dark_downward_diagonal'   => 'dkDnDiag',
        'dark_upward_diagonal'     => 'dkUpDiag',
        'wide_downward_diagonal'   => 'wdDnDiag',
        'wide_upward_diagonal'     => 'wdUpDiag',

        'light_vertical'           => 'ltVert',
        'light_horizontal'         => 'ltHorz',
        'narrow_vertical'          => 'narVert',
        'narrow_horizontal'        => 'narHorz',
        'dark_vertical'            => 'dkVert',
        'dark_horizontal'          => 'dkHorz',

        'dashed_downward_diagonal' => 'dashDnDiag',
        'dashed_upward_diagonal'   => 'dashUpDiag',
        'dashed_horizontal'        => 'dashHorz',
        'dashed_vertical'          => 'dashVert',
        'small_confetti'           => 'smConfetti',
        'large_confetti'           => 'lgConfetti',

        'zigzag'                   => 'zigZag',
        'wave'                     => 'wave',
        'diagonal_brick'           => 'diagBrick',
        'horizontal_brick'         => 'horzBrick',
        'weave'                    => 'weave',
        'plaid'                    => 'plaid',

        'divot'                    => 'divot',
        'dotted_grid'              => 'dotGrid',
        'dotted_diamond'           => 'dotDmnd',
        'shingle'                  => 'shingle',
        'trellis'                  => 'trellis',
        'sphere'                   => 'sphere',

        'small_grid'               => 'smGrid',
        'large_grid'               => 'lgGrid',
        'small_check'              => 'smCheck',
        'large_check'              => 'lgCheck',
        'outlined_diamond'         => 'openDmnd',
        'solid_diamond'            => 'solidDmnd'
      }.freeze

      #
      # Convert user defined legend properties to the structure required internally.
      #
      def legend_properties(params)
        legend = Writexlsx::Chart::Legend.new

        legend.position      = params[:position] || 'right'
        legend.delete_series = params[:delete_series]
        legend.font          = convert_font_args(params[:font])

        # Set the legend layout.
        legend.layout = layout_properties(params[:layout])

        # Turn off the legend.
        legend.position = 'none' if params[:none]

        # Set the line properties for the legend.
        line = line_properties(params[:line])

        # Allow 'border' as a synonym for 'line'.
        line = line_properties(params[:border]) if params[:border]

        # Set the fill properties for the legend.
        fill = fill_properties(params[:fill])

        # Set the pattern properties for the legend.
        pattern = pattern_properties(params[:pattern])

        # Set the gradient fill properties for the legend.
        gradient = gradient_properties(params[:gradient])

        # Pattern fill overrides solid fill.
        fill = nil if pattern

        # Gradient fill overrides solid and pattern fills.
        if gradient
          pattern = nil
          fill    = nil
        end

        # Set the legend layout.
        layout = layout_properties(params[:layout])

        legend.line     = line
        legend.fill     = fill
        legend.pattern  = pattern
        legend.gradient = gradient
        legend.layout   = layout

        @legend = legend
      end

      #
      # Convert user defined layout properties to the format required internally.
      #
      def layout_properties(args, is_text = false)
        return unless ptrue?(args)

        properties = is_text ? %i[x y] : %i[x y width height]

        # Check for valid properties.
        args.each_key do |key|
          raise "Property '#{key}' not allowed in layout options\n" unless properties.include?(key.to_sym)
        end

        # Set the layout properties
        layout = {}
        properties.each do |property|
          value = args[property]
          # Convert to the format used by Excel for easier testing.
          layout[property] = sprintf("%.17g", value)
        end

        layout
      end

      #
      # Convert user defined line properties to the structure required internally.
      #
      def line_properties(line) # :nodoc:
        line_fill_properties(line) do
          value_or_raise(dash_types, line[:dash_type], 'dash type')
        end
      end

      #
      # Convert user defined fill properties to the structure required internally.
      #
      def fill_properties(fill) # :nodoc:
        line_fill_properties(fill)
      end

      #
      # Convert user defined pattern properties to the structure required internally.
      #
      def pattern_properties(args) # :nodoc:
        return nil unless args
        # Check the pattern type is present.
        return nil unless args.has_key?(:pattern)
        # Check the foreground color is present.
        return nil unless args.has_key?(:fg_color)

        pattern = {}

        type = PATTERN_TYPES[args[:pattern]]
        raise "Unknown pattern type '#{args[:pattern]}'" unless type

        pattern[:pattern] = type
        pattern[:bg_color] = args[:bg_color] || '#FFFFFF'
        pattern[:fg_color] = args[:fg_color]

        pattern
      end

      def line_fill_properties(params)
        return { _defined: 0 } unless params

        ret = params.dup
        ret[:dash_type] = yield if block_given? && ret[:dash_type]
        ret[:_defined] = 1
        ret
      end

      def dash_types
        {
          solid:               'solid',
          round_dot:           'sysDot',
          square_dot:          'sysDash',
          dash:                'dash',
          dash_dot:            'dashDot',
          long_dash:           'lgDash',
          long_dash_dot:       'lgDashDot',
          long_dash_dot_dot:   'lgDashDotDot',
          dot:                 'dot',
          system_dash_dot:     'sysDashDot',
          system_dash_dot_dot: 'sysDashDotDot'
        }
      end

      def value_or_raise(hash, key, msg)
        raise "Unknown #{msg} '#{key}'" if hash[key.to_sym].nil?

        hash[key.to_sym]
      end

      def palette_color_from_index(index)
        # Adjust the colour index.
        idx = index - 8

        r, g, b = @palette[idx]
        sprintf("%02X%02X%02X", r, g, b)
      end

      #
      # Convert the user specified colour index or string to a rgb colour.
      #
      def color(color_code) # :nodoc:
        if color_code && color_code =~ /^#[0-9a-fA-F]{6}$/
          # Convert a HTML style #RRGGBB color.
          color_code.sub(/^#/, '').upcase
        else
          index = Format.color(color_code)
          raise "Unknown color '#{color_code}' used in chart formatting." unless index

          palette_color_from_index(index)
        end
      end

      #
      # Write the <a:solidFill> element.
      #
      def write_a_solid_fill(fill) # :nodoc:
        @writer.tag_elements('a:solidFill') do
          if fill[:color]
            # Write the a:srgbClr element.
            write_a_srgb_clr(color(fill[:color]), fill[:transparency])
          end
        end
      end

      #
      # Write the <a:srgbClr> element.
      #
      def write_a_srgb_clr(color, transparency = nil) # :nodoc:
        tag        = 'a:srgbClr'
        attributes = [['val', color]]

        if ptrue?(transparency)
          @writer.tag_elements(tag, attributes) do
            write_a_alpha(transparency)
          end
        else
          @writer.empty_tag(tag, attributes)
        end
      end
    end
  end
end
