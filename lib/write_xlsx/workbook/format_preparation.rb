# -*- coding: utf-8 -*-
# frozen_string_literal: true

module Writexlsx
  class Workbook
    module FormatPreparation
      private

      #
      # Prepare all of the format properties prior to passing them to Styles.rb.
      #
      def prepare_format_properties # :nodoc:
        # Separate format objects into XF and DXF formats.
        prepare_formats

        # Set the font index for the format objects.
        prepare_fonts

        # Set the number format index for the format objects.
        prepare_num_formats

        # Set the border index for the format objects.
        prepare_borders

        # Set the fill index for the format objects.
        prepare_fills
      end

      #
      # Iterate through the XF Format objects and separate them into XF and DXF
      # formats.
      #
      def prepare_formats # :nodoc:
        @formats.formats.each do |format|
          xf_index  = format.xf_index
          dxf_index = format.dxf_index

          @xf_formats[xf_index] = format   if xf_index
          @dxf_formats[dxf_index] = format if dxf_index
        end
      end

      #
      # Iterate through the XF Format objects and give them an index to non-default
      # font elements.
      #
      def prepare_fonts # :nodoc:
        fonts = {}

        @xf_formats.each { |format| format.set_font_info(fonts) }

        @font_count = fonts.size

        # For the DXF formats we only need to check if the properties have changed.
        @dxf_formats.each do |format|
          # The only font properties that can change for a DXF format are: color,
          # bold, italic, underline and strikethrough.
          format.has_dxf_font(true) if format.color? || format.bold? || format.italic? || format.underline? || format.strikeout?
        end
      end

      #
      # Iterate through the XF Format objects and give them an index to non-default
      # number format elements.
      #
      # User defined records start from index 0xA4.
      #
      def prepare_num_formats # :nodoc:
        num_formats        = []
        unique_num_formats = {}
        index              = 164

        (@xf_formats + @dxf_formats).each do |format|
          num_format = format.num_format

          # Check if num_format is an index to a built-in number format.
          # Also check for a string of zeros, which is a valid number format
          # string but would evaluate to zero.
          #
          if num_format.to_s =~ /^\d+$/ && num_format.to_s !~ /^0+\d/
            # Number format '0' is indexed as 1 in Excel.
            num_format = 1 if num_format == 0
            # Index to a built-in number format.
            format.num_format_index = num_format
            next
          elsif num_format.to_s == 'General'
            # The 'General' format has an number format index of 0.
            format.num_format_index = 0
            next
          end

          if unique_num_formats[num_format]
            # Number format has already been used.
            format.num_format_index = unique_num_formats[num_format]
          else
            # Add a new number format.
            unique_num_formats[num_format] = index
            format.num_format_index = index
            index += 1

            # Only store/increase number format count for XF formats
            # (not for DXF formats).
            num_formats << num_format if ptrue?(format.xf_index)
          end
        end

        @num_formats = num_formats
      end

      #
      # Iterate through the XF Format objects and give them an index to non-default
      # border elements.
      #
      def prepare_borders # :nodoc:
        borders = {}

        @xf_formats.each { |format| format.set_border_info(borders) }

        @border_count = borders.size

        # For the DXF formats we only need to check if the properties have changed.
        @dxf_formats.each do |format|
          key = format.get_border_key
          format.has_dxf_border(true) if key =~ /[^0:]/
        end
      end

      #
      # Iterate through the XF Format objects and give them an index to non-default
      # fill elements.
      #
      # The user defined fill properties start from 2 since there are 2 default
      # fills: patternType="none" and patternType="gray125".
      #
      def prepare_fills # :nodoc:
        fills = {}
        index = 2    # Start from 2. See above.

        # Add the default fills.
        fills['0:0:0']  = 0
        fills['17:0:0'] = 1

        # Store the DXF colors separately since them may be reversed below.
        @dxf_formats.each do |format|
          next unless format.pattern != 0 || format.bg_color != 0 || format.fg_color != 0

          format.has_dxf_fill(true)
          format.dxf_bg_color = format.bg_color
          format.dxf_fg_color = format.fg_color
        end

        @xf_formats.each do |format|
          # The following logical statements jointly take care of special cases
          # in relation to cell colours and patterns:
          # 1. For a solid fill (_pattern == 1) Excel reverses the role of
          #    foreground and background colours, and
          # 2. If the user specifies a foreground or background colour without
          #    a pattern they probably wanted a solid fill, so we fill in the
          #    defaults.
          #
          if format.pattern == 1 && ne_0?(format.bg_color) && ne_0?(format.fg_color)
            format.fg_color, format.bg_color = format.bg_color, format.fg_color
          elsif format.pattern <= 1 && ne_0?(format.bg_color) && eq_0?(format.fg_color)
            format.fg_color = format.bg_color
            format.bg_color = 0
            format.pattern  = 1
          elsif format.pattern <= 1 && eq_0?(format.bg_color) && ne_0?(format.fg_color)
            format.bg_color = 0
            format.pattern  = 1
          end

          key = format.get_fill_key

          if fills[key]
            # Fill has already been used.
            format.fill_index = fills[key]
            format.has_fill(false)
          else
            # This is a new fill.
            fills[key]        = index
            format.fill_index = index
            format.has_fill(true)
            index += 1
          end
        end

        @fill_count = index
      end

      def eq_0?(val)
        !ptrue?(val)
      end

      def ne_0?(val)
        !eq_0?(val)
      end
    end
  end
end
