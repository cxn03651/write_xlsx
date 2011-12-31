# -*- coding: utf-8 -*-
module Writexlsx
  class Format
    attr_reader :xf_index, :dxf_index, :num_format
    attr_reader :bold, :italic, :font_strikeout, :font_shadow, :font_outline
    attr_reader :underline, :font_script, :size, :theme, :font, :font_family, :hyperlink
    attr_reader :diag_type, :diag_color, :font_only, :color, :color_indexed
    attr_reader :left, :left_color, :right, :right_color, :top, :top_color, :bottom, :bottom_color
    attr_reader :font_scheme
    attr_accessor :font_index, :has_font, :has_dxf_font, :has_dxf_fill, :has_dxf_border
    attr_accessor :num_format_index, :border_index, :has_border
    attr_accessor :fill_index, :has_fill, :font_condense, :font_extend, :diag_border
    attr_accessor :bg_color, :fg_color, :pattern

    def initialize(xf_format_indices = {}, dxf_format_indices = {}, params = {})   # :nodoc:
      @xf_format_indices = xf_format_indices
      @dxf_format_indices = dxf_format_indices

      @xf_index       = nil
      @dxf_index      = nil

      @num_format     = 0
      @num_format_index = 0
      @font_index     = 0
      @has_font       = 0
      @has_dxf_font   = 0
      @font           = 'Calibri'
      @size           = 11
      @bold           = 0
      @italic         = 0
      @color          = 0x0
      @underline      = 0
      @font_strikeout = 0
      @font_outline   = 0
      @font_shadow    = 0
      @font_script    = 0
      @font_family    = 2
      @font_charset   = 0
      @font_scheme    = 'minor'
      @font_condense  = 0
      @font_extend    = 0
      @theme          = 0
      @hyperlink      = 0

      @hidden         = 0
      @locked         = 1

      @text_h_align   = 0
      @text_wrap      = 0
      @text_v_align   = 0
      @text_justlast  = 0
      @rotation       = 0

      @fg_color       = 0x00
      @bg_color       = 0x00
      @pattern        = 0
      @has_fill       = 0
      @has_dxf_fill   = 0
      @fill_index     = 0
      @fill_count     = 0

      @border_index   = 0
      @has_border     = 0
      @has_dxf_border = 0
      @border_count   = 0

      @bottom         = 0
      @bottom_color   = 0x0
      @diag_border    = 0
      @diag_color     = 0x0
      @diag_type      = 0
      @left           = 0
      @left_color     = 0x0
      @right          = 0
      @right_color    = 0x0
      @top            = 0
      @top_color      = 0x0

      @indent         = 0
      @shrink         = 0
      @merge_range    = 0
      @reading_order  = 0
      @just_distrib   = 0
      @color_indexed  = 0
      @font_only      = 0

      set_format_properties(params) unless params.empty?
    end

    #
    # :call-seq:
    #    set_format_properties( :bold => 1 [, :color => 'red'..] )
    #    set_format_properties( font [, shade, ..])
    #    set_format_properties( :bold => 1, font, ...)
    #      *) font  = { :color => 'red', :bold => 1 }
    #         shade = { :bg_color => 'green', :pattern => 1 }
    #
    # Convert hashes of properties to method calls.
    #
    # The properties of an existing Format object can be also be set by means
    # of set_format_properties():
    #
    #     format = workbook.add_format
    #     format.set_format_properties(:bold => 1, :color => 'red');
    #
    # However, this method is here mainly for legacy reasons. It is preferable
    # to set the properties in the format constructor:
    #
    #     format = workbook.add_format(:bold => 1, :color => 'red');
    #
    def set_format_properties(*properties)   # :nodoc:
      return if properties.empty?
      properties.each do |property|
        property.each do |key, value|
          # Strip leading "-" from Tk style properties e.g. "-color" => 'red'.
          key = key.sub(/^-/, '') if key.respond_to?(:to_str)

          # Create a sub to set the property.
          if value.respond_to?(:to_str) || !value.respond_to?(:+)
            s = "set_#{key}('#{value}')"
          else
            s = "set_#{key}(#{value})"
          end
          eval s
        end
      end
    end

    #
    # Return properties for an Style xf <alignment> sub-element.
    #
    def get_align_properties
      align = []    # Attributes to return

      # Check if any alignment options in the format have been changed.
      if @text_h_align != 0 || @text_v_align != 0 || @indent != 0 ||
         @rotation != 0 || @text_wrap != 0 || @shrink != 0 || @reading_order != 0
        changed = 1
      else
        return
      end

      # Indent is only allowed for horizontal left, right and distributed. If it
      # is defined for any other alignment or no alignment has been set then
      # default to left alignment.
      @text_h_align = 1 if @indent != 0 && ![1, 3, 7].include?(@text_h_align)

      # Check for properties that are mutually exclusive.
      @shrink       = 0 if @text_wrap != 0
      @shrink       = 0 if @text_h_align == 4    # Fill
      @shrink       = 0 if @text_h_align == 5    # Justify
      @shrink       = 0 if @text_h_align == 7    # Distributed
      @just_distrib = 0 if @text_h_align != 7    # Distributed
      @just_distrib = 0 if @indent != 0

      continuous = 'centerContinuous'

      align << 'horizontal' << 'left'        if @text_h_align == 1
      align << 'horizontal' << 'center'      if @text_h_align == 2
      align << 'horizontal' << 'right'       if @text_h_align == 3
      align << 'horizontal' << 'fill'        if @text_h_align == 4
      align << 'horizontal' << 'justify'     if @text_h_align == 5
      align << 'horizontal' << continuous    if @text_h_align == 6
      align << 'horizontal' << 'distributed' if @text_h_align == 7

      align << 'justifyLastLine' << 1 if @just_distrib != 0

      # Property 'vertical' => 'bottom' is a default. It sets applyAlignment
      # without an alignment sub-element.
      align << 'vertical' << 'top'         if @text_v_align == 1
      align << 'vertical' << 'center'      if @text_v_align == 2
      align << 'vertical' << 'justify'     if @text_v_align == 4
      align << 'vertical' << 'distributed' if @text_v_align == 5

      align << 'indent' <<       @indent   if @indent   != 0
      align << 'textRotation' << @rotation if @rotation != 0

      align << 'wrapText' <<     1 if @text_wrap != 0
      align << 'shrinkToFit' <<  1 if @shrink    != 0

      align << 'readingOrder' << 1 if @reading_order == 1
      align << 'readingOrder' << 2 if @reading_order == 2

      return changed, align
    end

    #
    # Return properties for an Excel XML <Protection> element.
    #
    def get_protection_properties
      attributes = []

      attributes << 'locked' << 0 if     @locked == 0
      attributes << 'hidden' << 1 unless @hidden == 0

      attributes.empty? ? nil : attributes
    end

    def set_bold(bold = 1)
      @bold = (bold && bold != 0) ? 1 : 0
    end

    def inspect
      to_s
    end

    #
    # Returns a unique hash key for the Format object.
    #
    def get_format_key
      [get_font_key, get_border_key, get_fill_key, @num_format].join(':')
    end

    #
    # Returns a unique hash key for a font. Used by Workbook.
    #
    def get_font_key
      [
        @bold,
        @color,
        @font_charset,
        @font_family,
        @font_outline,
        @font_script,
        @font_shadow,
        @font_strikeout,
        @font,
        @italic,
        @size,
        @underline
      ].join(':')
    end

    #
    # Returns a unique hash key for a border style. Used by Workbook.
    #
    def get_border_key
      [
        @bottom,
        @bottom_color,
        @diag_border,
        @diag_color,
        @diag_type,
        @left,
        @left_color,
        @right,
        @right_color,
        @top,
        @top_color
      ].join(':')
    end

    #
    # Returns a unique hash key for a fill style. Used by Workbook.
    #
    def get_fill_key
      [
        @pattern,
        @bg_color,
        @fg_color
      ].join(':')
    end

    #
    # Returns the index used by Worksheet->_XF()
    #
    def get_xf_index
      if @xf_index
        @xf_index
      else
        key = get_format_key
        indices_href = @xf_format_indices

        if indices_href[key]
          indices_href[key]
        else
          index = 1 + indices_href.keys.size
          indices_href[key] = index
          @xf_index = index
          index
        end
      end
    end

    #
    # Returns the index used by Worksheet->_XF()
    #
    def get_dxf_index
      if @dxf_index
          @dxf_index
      else
        key  = get_format_key
        indices = @dxf_format_indices

        if indices[key]
          indices[key]
        else
          index = indices.size
          indices[key] = index
          @dxf_index = index
          index
        end
      end
    end

    def get_color(color)
      Format.get_color(color)
    end

    #
    # Used in conjunction with the set_xxx_color methods to convert a color
    # string into a number. Color range is 0..63 but we will restrict it
    # to 8..63 to comply with Gnumeric. Colors 0..7 are repeated in 8..15.
    #
    def self.get_color(color)

      colors = {
        :aqua    => 0x0F,
        :cyan    => 0x0F,
        :black   => 0x08,
        :blue    => 0x0C,
        :brown   => 0x10,
        :magenta => 0x0E,
        :fuchsia => 0x0E,
        :gray    => 0x17,
        :grey    => 0x17,
        :green   => 0x11,
        :lime    => 0x0B,
        :navy    => 0x12,
        :orange  => 0x35,
        :pink    => 0x21,
        :purple  => 0x14,
        :red     => 0x0A,
        :silver  => 0x16,
        :white   => 0x09,
        :yellow  => 0x0D,
      }

      if color.respond_to?(:to_str)
        # Return RGB style colors for processing later.
        return color if color =~ /^#[0-9A-F]{6}$/i

        # Return the default color if undef,
        return 0x00 unless color

        # or the color string converted to an integer,
        return colors[color.downcase.to_sym] if colors[color.downcase.to_sym]

        # or the default color if string is unrecognised,
        return 0x00 if color =~ /\D/
      else
        # or an index < 8 mapped into the correct range,
        return color + 8 if color < 8

        # or the default color if arg is outside range,
        return 0x00 if color > 63

        # or an integer in the valid range
        return color
      end
    end

    #
    # Set cell alignment.
    #
    def set_align(location)
      return unless location             # No default

      location.downcase!

      set_text_h_align(1) if location == 'left'
      set_text_h_align(2) if location == 'centre'
      set_text_h_align(2) if location == 'center'
      set_text_h_align(3) if location == 'right'
      set_text_h_align(4) if location == 'fill'
      set_text_h_align(5) if location == 'justify'
      set_text_h_align(6) if location == 'center_across'
      set_text_h_align(6) if location == 'centre_across'
      set_text_h_align(6) if location == 'merge'              # Legacy.
      set_text_h_align(7) if location == 'distributed'
      set_text_h_align(7) if location == 'equal_space'        # S::PE.
      set_text_h_align(7) if location == 'justify_distributed'

      @just_distrib =   1 if location == 'justify_distributed'

      set_text_v_align(1) if location == 'top'
      set_text_v_align(2) if location == 'vcentre'
      set_text_v_align(2) if location == 'vcenter'
      set_text_v_align(3) if location == 'bottom'
      set_text_v_align(4) if location == 'vjustify'
      set_text_v_align(5) if location == 'vdistributed'
      set_text_v_align(5) if location == 'vequal_space'    # S::PE.
    end

    #
    # Set vertical cell alignment. This is required by the set_properties() method
    # to differentiate between the vertical and horizontal properties.
    #
    def set_valign(location)
      set_align(location)
    end

    #
    # Implements the Excel5 style "merge".
    #
    def set_center_across(flag = 1)
      set_text_h_align(6)
    end

    #
    # This was the way to implement a merge in Excel5. However it should have been
    # called "center_across" and not "merge".
    # This is now deprecated. Use set_center_across() or better merge_range().
    #
    def set_merge(merge = 1)
      set_text_h_align(6)
    end

    #
    # Set cells borders to the same style
    #
    def set_border(style)
      set_bottom(style)
      set_top(style)
      set_left(style)
      set_right(style)
    end

    #
    # Set cells border to the same color
    #
    def set_border_color(color)
      set_bottom_color(color)
      set_top_color(color)
      set_left_color(color)
      set_right_color(color)
    end

    #
    # Set the rotation angle of the text. An alignment property.
    #
    def set_rotation(rotation)
      if rotation == 270
        rotation = 255
      elsif rotation >= -90 || rotation <= 90
        rotation = -rotation + 90 if rotation < 0
      else
        raise "Rotation #{rotation} outside range: -90 <= angle <= 90"
        rotation = 0
      end

      @rotation = rotation
    end

    #
    # Set the properties for the hyperlink style. TODO. This doesn't currently
    # work. Fix it when styles are supported.
    #
    def set_hyperlink
      @hyperlink = 1

      set_underline(1)
      set_theme(10)
      set_align('top')
    end

    def method_missing(name, *args)  # :nodoc:
      method = "#{name}"

      # Check for a valid method names, i.e. "set_xxx_yyy".
      method =~ /set_(\w+)/ or raise "Unknown method: #{method}\n"

      # Match the attribute, i.e. "@xxx_yyy".
      attribute = "@#{$1}"

      # Check that the attribute exists
      # ........
      if method =~ /set\w+color$/    # for "set_property_color" methods
        value = get_color(args[0])
      else                            # for "set_xxx" methods
        value = args[0].nil? ? 1 : args[0]
      end
      if value.respond_to?(:to_str) || !value.respond_to?(:+)
        s = %Q!#{attribute} = "#{value.to_s}"!
      else
        s = %Q!#{attribute} =   #{value.to_s}!
      end
      eval s
    end
  end
end
