# -*- coding: utf-8 -*-
require 'write_xlsx/package/xml_writer_simple'
require 'write_xlsx/utility'

module Writexlsx
  module Package
    class Vml

      include Writexlsx::Utility

      def initialize
        @writer = Package::XMLWriterSimple.new
      end

      def set_xml_writer(filename)
        @writer.set_xml_writer(filename)
      end

      def assemble_xml_file(worksheet)
        return unless @writer

        write_xml_namespace

        # Write the o:shapelayout element.
        write_shapelayout(worksheet.vml_data_id)

        # Write the v:shapetype element.
        write_shapetype

        z_index = 1
        vml_shape_id = worksheet.vml_shape_id
        worksheet.comments_array.each do |comment|
          # Write the v:shape element.
          vml_shape_id += 1
          write_shape(vml_shape_id, z_index, comment)
          z_index += 1
        end

        @writer.end_tag('xml')
        @writer.crlf
        @writer.close
      end

      private

      #
      # Convert comment vertices from pixels to points.
      #
      def pixels_to_points(vertices)
        col_start, row_start, x1,    y1,
        col_end,   row_end,   x2,    y2,
        left,      top,       width, height  = vertices.flatten

        left   *= 0.75
        top    *= 0.75
        width  *= 0.75
        height *= 0.75

        [left, top, width, height]
      end

      #
      # Write the <xml> element. This is the root element of VML.
      #
      def write_xml_namespace
        schema  = 'urn:schemas-microsoft-com:'
        xmlns   = schema + 'vml'
        xmlns_o = schema + 'office:office'
        xmlns_x = schema + 'office:excel'

        attributes = [
          'xmlns:v', xmlns,
          'xmlns:o', xmlns_o,
          'xmlns:x', xmlns_x
        ]

        @writer.start_tag('xml', attributes)
      end

      #
      # Write the <o:shapelayout> element.
      #
      def write_shapelayout(data_id)
        ext     = 'edit'

        attributes = ['v:ext', ext]

        @writer.tag_elements('o:shapelayout', attributes) do
          # Write the o:idmap element.
          write_idmap(data_id)
        end
      end

      #
      # Write the <o:idmap> element.
      #
      def write_idmap(data_id)
        ext     = 'edit'

        attributes = [
          'v:ext', ext,
          'data',  data_id
        ]

        @writer.empty_tag('o:idmap', attributes)
      end

      #
      # Write the <v:shapetype> element.
      #
      def write_shapetype
        id        = '_x0000_t202'
        coordsize = '21600,21600'
        spt       = 202
        path      = 'm,l,21600r21600,l21600,xe'

        attributes = [
            'id',        id,
            'coordsize', coordsize,
            'o:spt',     spt,
            'path',      path
        ]

        @writer.tag_elements('v:shapetype', attributes) do
          # Write the v:stroke element.
          write_stroke
          # Write the v:path element.
          write_path('t', 'rect')
        end
      end

      #
      # Write the <v:stroke> element.
      #
      def write_stroke
        joinstyle = 'miter'

        attributes = ['joinstyle', joinstyle]

        @writer.empty_tag('v:stroke', attributes)
      end

      #
      # Write the <v:path> element.
      #
      def write_path(gradientshapeok, connecttype)
        attributes      = []

        attributes << 'gradientshapeok' << 't' if gradientshapeok
        attributes << 'o:connecttype' << connecttype

        @writer.empty_tag('v:path', attributes)
      end

      #
      # Write the <v:shape> element.
      #
      def write_shape(id, z_index, comment)
        type       = '#_x0000_t202'
        insetmode  = 'auto'
        visibility = 'hidden'

        # Set the shape index.
        id = '_x0000_s' + id.to_s

        left, top, width, height = pixels_to_points(comment.vertices)

        # Set the visibility.
        visibility = 'visible' if comment.visible != 0 && !comment.visible.nil?

        left_str    = float_to_str(left)
        top_str     = float_to_str(top)
        width_str   = float_to_str(width)
        height_str  = float_to_str(height)
        z_index_str = float_to_str(z_index)

        style =
            'position:absolute;' +
            'margin-left:'       +
            left_str + 'pt;'     +
            'margin-top:'        +
            top_str + 'pt;'      +
            'width:'             +
            width_str + 'pt;'    +
            'height:'            +
            height_str + 'pt;'   +
            'z-index:'           +
            z_index_str + ';'    +
            'visibility:'        +
            visibility


        attributes = [
            'id',          id,
            'type',        type,
            'style',       style,
            'fillcolor',   comment.color,
            'o:insetmode', insetmode
        ]

        @writer.tag_elements('v:shape', attributes) do
          # Write the v:fill element.
          write_fill
          # Write the v:shadow element.
          write_shadow
          # Write the v:path element.
          write_path(nil, 'none')
          # Write the v:textbox element.
          write_textbox
          # Write the x:ClientData element.
          write_client_data(comment)
        end
      end

      def float_to_str(float)
        return '' unless float
        if float == float.to_i
          float.to_i.to_s
        else
          float.to_s
        end
      end

      #
      # Write the <v:fill> element.
      #
      def write_fill
        color_2 = '#ffffe1'
        attributes = ['color2', color_2]

        @writer.empty_tag('v:fill', attributes)
      end

      #
      # Write the <v:shadow> element.
      #
      def write_shadow
        on       = 't'
        color    = 'black'
        obscured = 't'

        attributes = [
            'on',       on,
            'color',    color,
            'obscured', obscured
        ]

        @writer.empty_tag('v:shadow', attributes)
      end

      #
      # Write the <v:textbox> element.
      #
      def write_textbox
        style = 'mso-direction-alt:auto'

        attributes = ['style', style]

        @writer.tag_elements('v:textbox', attributes) do
          # Write the div element.
          write_div
        end
      end

      #
      # Write the <div> element.
      #
      def write_div
        style = 'text-align:left'
        attributes = ['style', style]

        @writer.tag_elements('div', attributes) { }
      end

      #
      # Write the <x:ClientData> element.
      #
      def write_client_data(comment)
        object_type = 'Note'

        attributes = ['ObjectType', object_type]

        @writer.tag_elements('x:ClientData', attributes) do
          # Write the x:MoveWithCells element.
          write_move_with_cells
          # Write the x:SizeWithCells element.
          write_size_with_cells
          # Write the x:Anchor element.
          write_anchor(comment.vertices)
          # Write the x:AutoFill element.
          write_auto_fill
          # Write the x:Row element.
          write_row(comment.row)
          # Write the x:Column element.
          write_column(comment.col)
          # Write the x:Visible element.
          write_visible if comment.visible != 0 && !comment.visible.nil?
        end
      end

      #
      # Write the <x:MoveWithCells> element.
      #
      def write_move_with_cells
        @writer.empty_tag('x:MoveWithCells')
      end

      #
      # Write the <x:SizeWithCells> element.
      #
      def write_size_with_cells
        @writer.empty_tag('x:SizeWithCells')
      end

      #
      # Write the <x:Visible> element.
      #
      def write_visible
        @writer.empty_tag('x:Visible')
      end

      #
      # Write the <x:Anchor> element.
      #
      def write_anchor(vertices)
        col_start, row_start, x1, y1, col_end, row_end, x2, y2 = vertices
        data = [col_start, x1, row_start, y1, col_end, x2, row_end, y2].join(', ')

        @writer.data_element('x:Anchor', data)
      end

      #
      # Write the <x:AutoFill> element.
      #
      def write_auto_fill
        data = 'False'

        @writer.data_element('x:AutoFill', data)
      end

      #
      # Write the <x:Row> element.
      #
      def write_row(data)
        @writer.data_element('x:Row', data)
      end

      #
      # Write the <x:Column> element.
      #
      def write_column(data)
        @writer.data_element('x:Column', data)
      end
    end
  end
end
