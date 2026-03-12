# -*- coding: utf-8 -*-
# frozen_string_literal: true

module Writexlsx
  class Chart
    class ChartArea
      include Writexlsx::Utility
      include Writexlsx::Gradient

      attr_reader :line, :fill, :pattern, :gradient, :layout

      def initialize(params = {})
        @layout = layout_properties(params[:layout])

        # Allow 'border' as a synonym for 'line'.
        border = params_to_border(params)

        # Set the line properties for the chartarea.
        @line = line_properties(border || params[:line])

        # Set the pattern properties for the series.
        @pattern = pattern_properties(params[:pattern])

        # Set the gradient fill properties for the series.
        @gradient = gradient_properties(params[:gradient])

        # Map deprecated Spreadsheet::WriteExcel fill colour.
        fill = params[:color] ? { color: params[:color] } : params[:fill]
        @fill = fill_properties(fill)

        # Pattern fill overrides solid fill.
        @fill = nil if ptrue?(@pattern)

        # Gradient fill overrides solid and pattern fills.
        if ptrue?(@gradient)
          @pattern = nil
          @fill    = nil
        end
      end

      private

      def params_to_border(params)
        line_weight  = params[:line_weight]

        # Map deprecated Spreadsheet::WriteExcel line_weight.
        border = params[:border]
        border = { width: swe_line_weight(line_weight) } if line_weight

        # Map deprecated Spreadsheet::WriteExcel line_pattern.
        if params[:line_pattern]
          pattern = swe_line_pattern(params[:line_pattern])
          if pattern == 'none'
            border = { none: 1 }
          else
            border[:dash_type] = pattern
          end
        end

        # Map deprecated Spreadsheet::WriteExcel line colour.
        border[:color] = params[:line_color] if params[:line_color]
        border
      end

      #
      # Get the Spreadsheet::WriteExcel line pattern for backward compatibility.
      #
      def swe_line_pattern(val)
        swe_line_pattern_hash[numeric_or_downcase(val)] || 'solid'
      end

      def swe_line_pattern_hash
        {
          0              => 'solid',
          1              => 'dash',
          2              => 'dot',
          3              => 'dash_dot',
          4              => 'long_dash_dot_dot',
          5              => 'none',
          6              => 'solid',
          7              => 'solid',
          8              => 'solid',
          'solid'        => 'solid',
          'dash'         => 'dash',
          'dot'          => 'dot',
          'dash-dot'     => 'dash_dot',
          'dash-dot-dot' => 'long_dash_dot_dot',
          'none'         => 'none',
          'dark-gray'    => 'solid',
          'medium-gray'  => 'solid',
          'light-gray'   => 'solid'
        }
      end

      #
      # Get the Spreadsheet::WriteExcel line weight for backward compatibility.
      #
      def swe_line_weight(val)
        swe_line_weight_hash[numeric_or_downcase(val)] || 1
      end

      def swe_line_weight_hash
        {
          1          => 0.25,
          2          => 1,
          3          => 2,
          4          => 3,
          'hairline' => 0.25,
          'narrow'   => 1,
          'medium'   => 2,
          'wide'     => 3
        }
      end

      def numeric_or_downcase(val)
        val.respond_to?(:coerce) ? val : val.downcase
      end
    end
  end
end
