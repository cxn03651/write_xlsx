# -*- coding: utf-8 -*-
require 'write_xlsx/package/xml_writer_simple'
require 'write_xlsx/utility'

module Writexlsx
  module Package
    class Styles

      include Writexlsx::Utility

      def initialize
        @writer = Package::XMLWriterSimple.new
        @xf_formats       = nil
        @palette          = []
        @font_count       = 0
        @num_format_count = 0
        @border_count     = 0
        @fill_count       = 0
        @custom_colors    = []
        @dxf_formats      = []
      end

      def set_xml_writer(filename)
        @writer.set_xml_writer(filename)
      end

      def assemble_xml_file
        write_xml_declaration
        write_style_sheet
        write_num_fmts
        write_fonts
        write_fills
        write_borders
        write_cell_style_xfs
        write_cell_xfs
        write_cell_styles
        write_dxfs
        write_table_styles
        write_colors
        @writer.end_tag('styleSheet')
        @writer.crlf
        @writer.close
      end

      #
      # Pass in the Format objects and other properties used to set the styles.
      #
      def set_style_properties(xf_formats, palette, font_count, num_format_count, border_count, fill_count, custom_colors, dxf_formats)
        @xf_formats       = xf_formats
        @palette          = palette
        @font_count       = font_count
        @num_format_count = num_format_count
        @border_count     = border_count
        @fill_count       = fill_count
        @custom_colors    = custom_colors
        @dxf_formats      = dxf_formats
      end

      #
      # Convert from an Excel internal colour index to a XML style #RRGGBB index
      # based on the default or user defined values in the Workbook palette.
      #
      def get_palette_color(index)
        palette = @palette

        # Handle colours in #XXXXXX RGB format.
        return "FF#{$1.upcase}" if index =~ /^#([0-9A-F]{6})$/i

        # Adjust the colour index.
        index -= 8

        # Palette is passed in from the Workbook class.
        rgb = @palette[index]

        sprintf("FF%02X%02X%02X", *rgb[0, 3])
      end

      #
      # Write the <styleSheet> element.
      #
      def write_style_sheet
        xmlns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

        attributes = ['xmlns', xmlns]

        @writer.start_tag('styleSheet', attributes)
      end

      #
      # Write the <numFmts> element.
      #
      def write_num_fmts
        count = @num_format_count

        return if count == 0

        attributes = ['count', count]

        @writer.start_tag('numFmts', attributes)

        # Write the numFmts elements.
        @xf_formats.each do |format|
          # Ignore built-in number formats, i.e., < 164.
          next unless format.num_format_index >= 164
          write_num_fmt(format.num_format_index, format.num_format)
        end

        @writer.end_tag('numFmts')
      end

      #
      # Write the <numFmt> element.
      #
      def write_num_fmt(num_fmt_id, format_code)
        format_codes = {
          0  => 'General',
          1  => '0',
          2  => '0.00',
          3  => '#,##0',
          4  => '#,##0.00',
          5  => '($#,##0_);($#,##0)',
          6  => '($#,##0_);[Red]($#,##0)',
          7  => '($#,##0.00_);($#,##0.00)',
          8  => '($#,##0.00_);[Red]($#,##0.00)',
          9  => '0%',
          10 => '0.00%',
          11 => '0.00E+00',
          12 => '# ?/?',
          13 => '# ??/??',
          14 => 'm/d/yy',
          15 => 'd-mmm-yy',
          16 => 'd-mmm',
          17 => 'mmm-yy',
          18 => 'h:mm AM/PM',
          19 => 'h:mm:ss AM/PM',
          20 => 'h:mm',
          21 => 'h:mm:ss',
          22 => 'm/d/yy h:mm',
          37 => '(#,##0_);(#,##0)',
          38 => '(#,##0_);[Red](#,##0)',
          39 => '(#,##0.00_);(#,##0.00)',
          40 => '(#,##0.00_);[Red](#,##0.00)',
          41 => '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)',
          42 => '_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)',
          43 => '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)',
          44 => '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)',
          45 => 'mm:ss',
          46 => '[h]:mm:ss',
          47 => 'mm:ss.0',
          48 => '##0.0E+0',
          49 => '@'
        }

        # Set the format code for built-in number formats.
        if num_fmt_id < 164
          if format_codes[num_fmt_id]
            format_code = format_codes[num_fmt_id]
          else
            format_code = 'General'
          end
        end

        attributes = [
          'numFmtId',   num_fmt_id,
          'formatCode', format_code
        ]

        @writer.empty_tag('numFmt', attributes)
      end

      #
      # Write the <fonts> element.
      #
      def write_fonts
        count = @font_count

        attributes = ['count', count]

        @writer.start_tag('fonts', attributes)

        # Write the font elements for format objects that have them.
        @xf_formats.each { |format| write_font(format) unless format.has_font == 0 }

        @writer.end_tag('fonts')
      end

      #
      # Write the <font> element.
      #
      def write_font(format, dxf_format = nil)
        @writer.start_tag('font')

        # The condense and extend elements are mainly used in dxf formats.
        write_condense unless format.font_condense == 0
        write_extend   unless format.font_extend   == 0

        @writer.empty_tag('b')       if format.bold?
        @writer.empty_tag('i')       if format.italic?
        @writer.empty_tag('strike')  if format.strikeout?
        @writer.empty_tag('outline') if format.outline?
        @writer.empty_tag('shadow')  if format.shadow?

        # Handle the underline variants.
        write_underline( format.underline ) if format.underline?

        write_vert_align('superscript') if format.font_script == 1
        write_vert_align('subscript')   if format.font_script == 2

        @writer.empty_tag('sz', ['val', format.size]) if !dxf_format

        theme = format.theme
        if theme != 0
          write_color('theme', theme)
        elsif format.color_indexed != 0
          write_color('indexed', format.color_indexed)
        elsif format.color != 0
          color = get_palette_color(format.color)
          write_color('rgb', color)
        elsif !dxf_format
          write_color('theme', 1)
        end

        if !dxf_format
          @writer.empty_tag('name',   ['val', format.font])
          @writer.empty_tag('family', ['val', format.font_family])

          if format.font == 'Calibri' && format.hyperlink == 0
            @writer.empty_tag('scheme', ['val', format.font_scheme])
          end
        end

        @writer.end_tag('font')
      end

      #
      # _write_underline()
      #
      # Write the underline font element.
      #
      def write_underline(underline)
        # Handle the underline variants.
        if underline == 2
          attributes = ['val', 'double']
        elsif underline == 33
          attributes = ['val', 'singleAccounting']
        elsif underline == 34
          attributes = ['val', 'doubleAccounting']
        else
          attributes = []    # Default to single underline.
        end

        @writer.empty_tag('u', attributes)
      end

      #
      # Write the <color> element.
      #
      def write_color(name, value)
        attributes = [name, value]

        @writer.empty_tag('color', attributes)
      end

      #
      # Write the <fills> element.
      #
      def write_fills
        count = @fill_count

        attributes = ['count', count]

        @writer.start_tag('fills', attributes)

        # Write the default fill element.
        write_default_fill('none')
        write_default_fill('gray125')

        # Write the fill elements for format objects that have them.
        @xf_formats.each do |format|
          write_fill(format) unless format.has_fill == 0
        end

        @writer.end_tag( 'fills' )
      end

      #
      # Write the <fill> element for the default fills.
      #
      def write_default_fill(pattern_type)
        @writer.start_tag('fill')
        @writer.empty_tag('patternFill', ['patternType', pattern_type])
        @writer.end_tag('fill')
      end

      #
      # Write the <fill> element.
      #
      def write_fill(format, dxf_format = nil)
        pattern    = format.pattern
        bg_color   = format.bg_color
        fg_color   = format.fg_color

        patterns = %w(
          none
          solid
          mediumGray
          darkGray
          lightGray
          darkHorizontal
          darkVertical
          darkDown
          darkUp
          darkGrid
          darkTrellis
          lightHorizontal
          lightVertical
          lightDown
          lightUp
          lightGrid
          lightTrellis
          gray125
          gray0625
        )

        @writer.start_tag('fill' )

        # The "none" pattern is handled differently for dxf formats.
        if dxf_format && format.pattern <= 1
          @writer.start_tag('patternFill')
        else
          @writer.start_tag('patternFill', ['patternType', patterns[format.pattern]])
        end

        unless fg_color == 0
          fg_color = get_palette_color(fg_color)
          @writer.empty_tag('fgColor', ['rgb', fg_color])
        end

        if bg_color != 0
          bg_color = get_palette_color(bg_color)
          @writer.empty_tag('bgColor', ['rgb', bg_color])
        else
          @writer.empty_tag('bgColor', ['indexed', 64]) if !dxf_format
        end

        @writer.end_tag('patternFill')
        @writer.end_tag('fill')
      end

      #
      # Write the <borders> element.
      #
      def write_borders
        count = @border_count

        attributes = ['count', count]

        @writer.start_tag('borders', attributes)

        # Write the border elements for format objects that have them.
        @xf_formats.each do |format|
          write_border(format) unless format.has_border == 0
        end

        @writer.end_tag('borders')
      end

      #
      # Write the <border> element.
      #
      def write_border(format, dxf_format = nil)
        attributes = []

        # Diagonal borders add attributes to the <border> element.
        if format.diag_type == 1
          attributes << 'diagonalUp'   << 1
        elsif format.diag_type == 2
          attributes << 'diagonalDown' << 1
        elsif format.diag_type == 3
          attributes << 'diagonalUp'   << 1
          attributes << 'diagonalDown' << 1
        end

        # Ensure that a default diag border is set if the diag type is set.
        format.diag_border = 1 if format.diag_type != 0 && format.diag_border == 0

        # Write the start border tag.
        @writer.start_tag('border', attributes)

        # Write the <border> sub elements.
        write_sub_border('left',   format.left,   format.left_color)
        write_sub_border('right',  format.right,  format.right_color)
        write_sub_border('top',    format.top,    format.top_color)
        write_sub_border('bottom', format.bottom, format.bottom_color)

        # Condition DXF formats don't allow diagonal borders
        if !dxf_format
          write_sub_border('diagonal', format.diag_border, format.diag_color)
        end

        if dxf_format
          write_sub_border('vertical')
          write_sub_border('horizontal')
        end

        @writer.end_tag('border')
      end

      #
      # Write the <border> sub elements such as <right>, <top>, etc.
      #
      def write_sub_border(type, style = 0, color = nil)
        if style == 0
          @writer.empty_tag(type)
          return
        end

        border_styles = %w(
          none
          thin
          medium
          dashed
          dotted
          thick
          double
          hair
          mediumDashed
          dashDot
          mediumDashDot
          dashDotDot
          mediumDashDotDot
          slantDashDot
        )

        attributes = [:style, border_styles[style]]

        @writer.start_tag(type, attributes)

        if color != 0
          color = get_palette_color(color)
          @writer.empty_tag('color', ['rgb', color])
        else
          @writer.empty_tag('color', ['auto', 1])
        end

        @writer.end_tag(type)
      end

      #
      # Write the <cellStyleXfs> element.
      #
      def write_cell_style_xfs
        count = 1

        attributes = ['count', count]

        @writer.start_tag('cellStyleXfs', attributes)

        # Write the style_xf element.
        write_style_xf

        @writer.end_tag('cellStyleXfs')
      end

      #
      # Write the <cellXfs> element.
      #
      def write_cell_xfs
        formats = @xf_formats

        # Workaround for when the last format is used for the comment font
        # and shouldn't be used for cellXfs.
        last_format =   formats[-1]

        formats.pop if last_format && last_format.font_only != 0

        attributes = ['count', formats.size]

        @writer.start_tag('cellXfs', attributes)

        # Write the xf elements.
        formats.each { |format| write_xf(format) }

        @writer.end_tag('cellXfs')
      end

      #
      # Write the style <xf> element.
      #
      def write_style_xf
        num_fmt_id = 0
        font_id    = 0
        fill_id    = 0
        border_id  = 0

        attributes = [
          'numFmtId', num_fmt_id,
          'fontId',   font_id,
          'fillId',   fill_id,
          'borderId', border_id
        ]

        @writer.empty_tag('xf', attributes)
      end

      private

      #
      # Write the <xf> element.
      #
      def write_xf(format)
        has_align   = false
        has_protect = false

        attributes = [
            'numFmtId', format.num_format_index,
            'fontId'  , format.font_index,
            'fillId'  , format.fill_index,
            'borderId', format.border_index,
            'xfId'    , 0
        ]

        attributes << 'applyNumberFormat' << 1 if format.num_format_index > 0

        # Add applyFont attribute if XF format uses a font element.
        attributes << 'applyFont' << 1 if format.font_index > 0

        # Add applyFill attribute if XF format uses a fill element.
        attributes << 'applyFill' << 1 if format.fill_index > 0

        # Add applyBorder attribute if XF format uses a border element.
        attributes << 'applyBorder' << 1 if format.border_index > 0

        # Check if XF format has alignment properties set.
        apply_align, align = format.get_align_properties

        # Check if an alignment sub-element should be written.
        has_align = true if apply_align && !align.empty?

        # We can also have applyAlignment without a sub-element.
        attributes << 'applyAlignment' << 1 if apply_align

        # Check for cell protection properties.
        protection = format.get_protection_properties

        if protection
          attributes << 'applyProtection' << 1
          has_protect = true
        end

        # Write XF with sub-elements if required.
        if has_align || has_protect
          @writer.start_tag('xf', attributes)
          @writer.empty_tag('alignment',  align)      if has_align
          @writer.empty_tag('protection', protection) if has_protect
          @writer.end_tag('xf')
        else
          @writer.empty_tag('xf', attributes)
        end
      end

      #
      # Write the <cellStyles> element.
      #
      def write_cell_styles
        count = 1

        attributes = ['count', count]

        @writer.start_tag('cellStyles', attributes)

        # Write the cellStyle element.
        write_cell_style

        @writer.end_tag('cellStyles')
      end

      #
      # Write the <cellStyle> element.
      #
      def write_cell_style
        name       = 'Normal'
        xf_id      = 0
        builtin_id = 0

        attributes = [
            'name',      name,
            'xfId',      xf_id,
            'builtinId', builtin_id
        ]

        @writer.empty_tag('cellStyle', attributes)
      end

      #
      # Write the <dxfs> element.
      #
      def write_dxfs
        formats = @dxf_formats

        count = formats.size

        attributes = ['count', count]

        if !formats.empty?
          @writer.start_tag('dxfs', attributes)

          # Write the font elements for format objects that have them.
          @dxf_formats.each do |format|
            @writer.start_tag('dxf')
            write_font(format, 1) unless format.has_dxf_font == 0

            if format.num_format_index != 0
              write_num_fmt(format.num_format_index, format.num_format)
            end

            write_fill(format, 1)    if format.has_dxf_fill   != 0
            write_border(format, 1)  if format.has_dxf_border != 0
            @writer.end_tag('dxf')
          end

          @writer.end_tag('dxfs')
        else
          @writer.empty_tag('dxfs', attributes)
        end
      end

      #
      # Write the <tableStyles> element.
      #
      def write_table_styles
        count               = 0
        default_table_style = 'TableStyleMedium9'
        default_pivot_style = 'PivotStyleLight16'

        attributes = [
            'count',             count,
            'defaultTableStyle', default_table_style,
            'defaultPivotStyle', default_pivot_style
        ]

        @writer.empty_tag('tableStyles', attributes)
      end

      #
      # Write the <colors> element.
      #
      def write_colors
        custom_colors = @custom_colors

        return if @custom_colors.empty?

        @writer.start_tag( 'colors' )
        write_mru_colors(@custom_colors)
        @writer.end_tag('colors')
      end

      #
      # Write the <mruColors> element for the most recently used colours.
      #
      def write_mru_colors(*args)
        custom_colors = args

        # Limit the mruColors to the last 10.
        count = custom_colors.size
        custom_colors = custom_colors[-10, 10] if count > 10

        @writer.start_tag('mruColors')

        # Write the custom colors in reverse order.
        @custom_colors.reverse.each do |color|
          write_color('rgb', color)
        end

        @writer.end_tag('mruColors')
      end

      def write_xml_declaration
        @writer.xml_decl
      end

      #
      # Write the <condense> element.
      #
      def write_condense
        val  = 0

        attributes = ['val', val]

        @writer.empty_tag('condense', attributes)
      end

      #
      # Write the <extend> element.
      #
      def write_extend
        val  = 0

        attributes = ['val', val]

        @writer.empty_tag('extend', attributes)
      end
    end
  end
end
