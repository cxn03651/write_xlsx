# -*- coding: utf-8 -*-
# frozen_string_literal: true

module Writexlsx
  module Utility
    module Drawing
      #
      # Convert vertices from pixels to points.
      #
      def pixels_to_points(vertices)
        _col_start, _row_start, _x1,    _y1,
        _col_end,   _row_end,   _x2,    _y2,
        left,      top,       width, height  = vertices.flatten

        left   *= 0.75
        top    *= 0.75
        width  *= 0.75
        height *= 0.75

        [left, top, width, height]
      end

      def v_shape_attributes_base(id)
        [
          ['id',    "_x0000_s#{id}"],
          ['type',  type]
        ]
      end

      def v_shape_style_base(z_index, vertices)
        left, top, width, height = pixels_to_points(vertices)

        left_str    = float_to_str(left)
        top_str     = float_to_str(top)
        width_str   = float_to_str(width)
        height_str  = float_to_str(height)
        z_index_str = float_to_str(z_index)

        shape_style_base(left_str, top_str, width_str, height_str, z_index_str)
      end

      def shape_style_base(left_str, top_str, width_str, height_str, z_index_str)
        [
          'position:absolute;',
          'margin-left:',
          left_str, 'pt;',
          'margin-top:',
          top_str, 'pt;',
          'width:',
          width_str, 'pt;',
          'height:',
          height_str, 'pt;',
          'z-index:',
          z_index_str, ';'
        ]
      end

      #
      # Write the <v:fill> element.
      #
      def write_fill
        @writer.empty_tag('v:fill', fill_attributes)
      end

      #
      # Write the <v:path> element.
      #
      def write_comment_path(gradientshapeok, connecttype)
        attributes = []

        attributes << %w[gradientshapeok t] if gradientshapeok
        attributes << ['o:connecttype', connecttype]

        @writer.empty_tag('v:path', attributes)
      end

      #
      # Write the <x:Anchor> element.
      #
      def write_anchor
        col_start, row_start, x1, y1, col_end, row_end, x2, y2 = @vertices
        data = [col_start, x1, row_start, y1, col_end, x2, row_end, y2].join(', ')

        @writer.data_element('x:Anchor', data)
      end

      #
      # Write the <x:AutoFill> element.
      #
      def write_auto_fill
        @writer.data_element('x:AutoFill', 'False')
      end

      #
      # Write the <div> element.
      #
      def write_div(align, font = nil)
        style = "text-align:#{align}"
        attributes = [['style', style]]

        @writer.tag_elements('div', attributes) do
          if font
            # Write the font element.
            write_font(font)
          end
        end
      end

      #
      # Write the <font> element.
      #
      def write_font(font)
        caption = font[:_caption]
        face    = 'Calibri'
        size    = 220
        color   = '#000000'

        attributes = [
          ['face',  face],
          ['size',  size],
          ['color', color]
        ]
        @writer.data_element('font', caption, attributes)
      end

      #
      # Write the <v:stroke> element.
      #
      def write_stroke
        attributes = [%w[joinstyle miter]]

        @writer.empty_tag('v:stroke', attributes)
      end
    end
  end
end
