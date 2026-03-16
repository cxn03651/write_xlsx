# -*- coding: utf-8 -*-
# frozen_string_literal: true

require 'write_xlsx/utility/common'
require 'write_xlsx/utility/xml_primitives'
require 'write_xlsx/format/fill_style'
require 'write_xlsx/format/fill_state'
require 'write_xlsx/format/border_style'
require 'write_xlsx/format/border_state'
require 'write_xlsx/format/font_style'
require 'write_xlsx/format/font_state'
require 'write_xlsx/format/alignment_style'
require 'write_xlsx/format/alignment_state'
require 'write_xlsx/format/protection_style'
require 'write_xlsx/format/protection_state'
require 'write_xlsx/format/number_format_style'
require 'write_xlsx/format/number_format_state'
require 'write_xlsx/format/format_state'

module Writexlsx
  #
  # A cell format in the workbook.
  #
  # Format represents the formatting properties associated with worksheet
  # cells. These properties are written to the workbook style tables
  # (styles.xml) and referenced by worksheet cell records.
  #
  # Responsibilities of this class include:
  #
  # * storing formatting properties such as font, fill, border, alignment,
  #   number format, and protection settings
  # * exposing the public `set_*` formatting API used by Workbook and Worksheet
  # * delegating grouped property access through style facade objects such as
  #   FillStyle, BorderStyle, FontStyle, AlignmentStyle, ProtectionStyle,
  #   and NumberFormatStyle
  # * generating unique format keys used by Workbook to deduplicate style
  #   records
  # * providing helper methods used when writing styles.xml
  #
  # Format instances are created by Workbook#add_format and reused across
  # worksheets through their computed format keys.
  #
  class Format
    include Writexlsx::Utility::Common
    include Writexlsx::Utility::XmlPrimitives

    ###########################################################################
    #
    # Lifecycle
    #
    ###########################################################################

    def initialize(formats, params = {}) # :nodoc:
      @formats = formats
      @state   = FormatState.new

      set_format_properties(params) unless params.empty?
    end

    #
    # Copy the attributes of another Format object.
    #
    def copy(other)
      reserve = %i[
        @xdf_format_indices
        @palette
      ]

      (instance_variables - reserve).each do |v|
        value = other.instance_variable_get(v)
        value = value.dup if v == :@state && !value.nil?
        instance_variable_set(v, value)
      end
    end

    ###########################################################################
    #
    # Style facade accessors
    #
    ###########################################################################

    def state
      @state
    end

    def fill_style
      FillStyle.new(self)
    end

    def border_style
      BorderStyle.new(self)
    end

    def font_style
      FontStyle.new(self)
    end

    def alignment_style
      AlignmentStyle.new(self)
    end

    def protection_style
      ProtectionStyle.new(self)
    end

    def number_format_style
      NumberFormatStyle.new(self)
    end

    ###########################################################################
    #
    # Explicit format property accessors
    #
    ###########################################################################

    #
    # Workbook indexes
    #
    def xf_index
      state.xf_index
    end

    def xf_index=(value)
      state.xf_index = value
    end

    def dxf_index
      state.dxf_index
    end

    def dxf_index=(value)
      state.dxf_index = value
    end

    def xf_id
      state.xf_id
    end

    def xf_id=(value)
      state.xf_id = value
    end

    #
    # Miscellaneous flags and compatibility ivars
    #
    def has_fill
      @has_fill
    end

    def has_fill=(value)
      @has_fill = value
    end

    def quote_prefix
      @quote_prefix
    end

    def quote_prefix=(value)
      @quote_prefix = value
    end

    def dxf_fg_color
      @dxf_fg_color
    end

    def dxf_fg_color=(value)
      @dxf_fg_color = value
    end

    def dxf_bg_color
      @dxf_bg_color
    end

    def dxf_bg_color=(value)
      @dxf_bg_color = value
    end

    #
    # Number format properties
    #
    def num_format
      number_format_style.format_code
    end

    def num_format=(value)
      number_format_style.format_code = value
    end

    def num_format_index
      number_format_style.index
    end

    def num_format_index=(value)
      number_format_style.index = value.to_i
    end

    #
    # Font properties
    #
    def font
      font_style.name
    end

    def font=(value)
      font_style.name = value
    end

    def size
      font_style.size
    end

    def size=(value)
      font_style.size = value
    end

    def font_color
      font_style.color
    end

    def font_color=(value)
      font_style.color = value
    end

    def bold
      font_style.bold
    end

    def bold=(value)
      font_style.bold = value
    end

    def italic
      font_style.italic
    end

    def italic=(value)
      font_style.italic = value
    end

    def underline
      font_style.underline
    end

    def underline=(value)
      font_style.underline = value
    end

    def font_strikeout
      font_style.strikeout
    end

    def font_strikeout=(value)
      font_style.strikeout = value
    end

    def font_script
      font_style.script
    end

    def font_script=(value)
      font_style.script = value
    end

    def font_outline
      font_style.outline
    end

    def font_outline=(value)
      font_style.outline = value
    end

    def font_shadow
      font_style.shadow
    end

    def font_shadow=(value)
      font_style.shadow = value
    end

    def font_charset
      font_style.charset
    end

    def font_charset=(value)
      font_style.charset = value
    end

    def font_family
      font_style.family
    end

    def font_family=(value)
      font_style.family = value
    end

    def font_scheme
      font_style.scheme
    end

    def font_scheme=(value)
      font_style.scheme = value
    end

    def font_condense
      font_style.condense
    end

    def font_condense=(value)
      font_style.condense = value
    end

    def font_extend
      font_style.extend
    end

    def font_extend=(value)
      font_style.extend = value
    end

    def color_indexed
      font_style.color_indexed
    end

    def color_indexed=(value)
      font_style.color_indexed = value
    end

    def theme
      font_style.theme
    end

    def theme=(value)
      font_style.theme = value
    end

    def hyperlink
      font_style.hyperlink
    end

    def hyperlink=(value)
      font_style.hyperlink = value
    end

    def font_index
      font_style.index
    end

    def font_index=(value)
      font_style.index = value
    end

    #
    # Protection properties
    #
    def locked
      protection_style.locked
    end

    def locked=(value)
      protection_style.locked = value
    end

    def hidden
      protection_style.hidden
    end

    def hidden=(value)
      protection_style.hidden = value
    end

    #
    # Alignment properties
    #
    def text_h_align
      alignment_style.horizontal
    end

    def text_h_align=(value)
      alignment_style.horizontal = value
    end

    def text_wrap
      alignment_style.wrap
    end

    def text_wrap=(value)
      alignment_style.wrap = value
    end

    def text_v_align
      alignment_style.vertical
    end

    def text_v_align=(value)
      alignment_style.vertical = value
    end

    def text_justlast
      alignment_style.justlast
    end

    def text_justlast=(value)
      alignment_style.justlast = value
    end

    def rotation
      alignment_style.rotation
    end

    def rotation=(value)
      alignment_style.rotation = value
    end

    def indent
      alignment_style.indent
    end

    def indent=(value)
      alignment_style.indent = value
    end

    def shrink
      alignment_style.shrink
    end

    def shrink=(value)
      alignment_style.shrink = value
    end

    def merge_range
      alignment_style.merge_range
    end

    def merge_range=(value)
      alignment_style.merge_range = value
    end

    def reading_order
      alignment_style.reading_order
    end

    def reading_order=(value)
      alignment_style.reading_order = value
    end

    def just_distrib
      alignment_style.just_distrib
    end

    def just_distrib=(value)
      alignment_style.just_distrib = value
    end

    #
    # Fill properties
    #
    def fg_color
      fill_style.fg_color
    end

    def bg_color
      fill_style.bg_color
    end

    def pattern
      fill_style.pattern
    end

    def fill_index
      fill_style.index
    end

    def fill_index=(value)
      fill_style.index = value
    end

    def fill_count
      fill_style.count
    end

    def fill_count=(value)
      fill_style.count = value
    end

    #
    # Border properties
    #
    def border_index
      border_style.index
    end

    def border_index=(value)
      border_style.index = value
    end

    def border_count
      border_style.count
    end

    def border_count=(value)
      border_style.count = value
    end

    def left
      border_style.left
    end

    def left=(value)
      border_style.left = value
    end

    def left_color
      border_style.left_color
    end

    def left_color=(value)
      border_style.left_color = value
    end

    def right
      border_style.right
    end

    def right=(value)
      border_style.right = value
    end

    def right_color
      border_style.right_color
    end

    def right_color=(value)
      border_style.right_color = value
    end

    def top
      border_style.top
    end

    def top=(value)
      border_style.top = value
    end

    def top_color
      border_style.top_color
    end

    def top_color=(value)
      border_style.top_color = value
    end

    def bottom
      border_style.bottom
    end

    def bottom=(value)
      border_style.bottom = value
    end

    def bottom_color
      border_style.bottom_color
    end

    def bottom_color=(value)
      border_style.bottom_color = value
    end

    def diag_border
      border_style.diag_border
    end

    def diag_border=(value)
      border_style.diag_border = value
    end

    def diag_color
      border_style.diag_color
    end

    def diag_color=(value)
      border_style.diag_color = value
    end

    def diag_type
      border_style.diag_type
    end

    def diag_type=(value)
      border_style.diag_type = value
    end

    ###########################################################################
    #
    # Public formatting property API
    #
    ###########################################################################

    def set_format_properties(*properties) # :nodoc:
      return if properties.empty?

      properties.each do |property|
        property.each do |key, value|
          send("set_#{key}", value)
        end
      end
    end

    #
    # Font setters
    #
    def set_font(value)
      self.font = normalize_format_property_value(value)
    end

    def set_font_family(value)
      self.font_family = value
    end

    def set_font_charset(value)
      self.font_charset = value
    end

    def set_font_condense(value)
      self.font_condense = value
    end

    def set_size(value)
      self.size = value
    end

    def set_color(value)
      self.font_color = color(normalize_format_property_value(value))
    end

    def set_bold(weight = 1)
      self.bold = weight
    end

    def set_italic(value = 1)
      self.italic = value
    end

    def set_underline(value = 1)
      self.underline = value
    end

    def set_font_strikeout(value = 1)
      self.font_strikeout = value
    end

    def set_font_script(value)
      self.font_script = value
    end

    def set_font_outline(value = 1)
      self.font_outline = value
    end

    def set_font_shadow(value = 1)
      self.font_shadow = value
    end

    def set_font_extend(value = 1)
      self.font_extend = value
    end

    def set_color_indexed(value = 1)
      self.color_indexed = value
    end

    def set_theme(value = 1)
      self.theme = value
    end

    #
    # Alignment setters
    #
    def set_align(location)
      return unless location

      location = location.downcase

      case location
      when 'left'                         then set_text_h_align(1)
      when 'centre', 'center'             then set_text_h_align(2)
      when 'right'                        then set_text_h_align(3)
      when 'fill'                         then set_text_h_align(4)
      when 'justify'                      then set_text_h_align(5)
      when 'center_across', 'centre_across', 'merge'
        set_text_h_align(6)
      when 'distributed', 'equal_space', 'justify_distributed'
        set_text_h_align(7)
      when 'top'                          then set_text_v_align(1)
      when 'vcentre', 'vcenter'           then set_text_v_align(2)
      when 'bottom'                       then set_text_v_align(3)
      when 'vjustify'                     then set_text_v_align(4)
      when 'vdistributed', 'vequal_space' then set_text_v_align(5)
      end

      self.just_distrib = 1 if location == 'justify_distributed'
    end

    def set_valign(location)
      set_align(location)
    end

    def set_center_across(_flag = 1)
      set_text_h_align(6)
    end

    def set_merge(_merge = 1)
      set_text_h_align(6)
    end

    def set_rotation(rotation)
      if rotation == 270
        rotation = 255
      elsif rotation.between?(-90, 90)
        rotation = -rotation + 90 if rotation < 0
      else
        raise "Rotation #{rotation} outside range: -90 <= angle <= 90"
      end

      self.rotation = rotation
    end

    def set_text_h_align(value)
      self.text_h_align = value
    end

    def set_text_v_align(value)
      self.text_v_align = value
    end

    def set_text_wrap(value = 1)
      self.text_wrap = value
    end

    def set_text_justlast(value = 1)
      self.text_justlast = value
    end

    def set_indent(value = 1)
      self.indent = value
    end

    def set_shrink(value = 1)
      self.shrink = value
    end

    #
    # Fill setters
    #
    def set_fg_color(value)
      fill_style.fg_color = color(normalize_format_property_value(value))
    end

    def set_bg_color(value)
      fill_style.bg_color = color(normalize_format_property_value(value))
    end

    def set_pattern(value)
      fill_style.pattern = normalize_format_property_value(value)
    end

    def set_fill_index(value)
      fill_style.index = normalize_format_property_value(value)
    end

    def set_fill_count(value)
      fill_style.count = normalize_format_property_value(value)
    end

    def set_has_fill(value)
      self.has_fill = value
    end

    #
    # Border setters
    #
    def set_border(value)
      set_bottom(value)
      set_top(value)
      set_left(value)
      set_right(value)
    end

    def set_border_color(value)
      color_value = color(normalize_format_property_value(value))
      self.bottom_color = color_value
      self.top_color    = color_value
      self.left_color   = color_value
      self.right_color  = color_value
    end

    def set_left(value)
      self.left = value
    end

    def set_right(value)
      self.right = value
    end

    def set_top(value)
      self.top = value
    end

    def set_bottom(value)
      self.bottom = value
    end

    def set_diag_border(value)
      self.diag_border = value
    end

    def set_diag_type(value)
      self.diag_type = value
    end

    def set_left_color(value)
      self.left_color = color(normalize_format_property_value(value))
    end

    def set_right_color(value)
      self.right_color = color(normalize_format_property_value(value))
    end

    def set_top_color(value)
      self.top_color = color(normalize_format_property_value(value))
    end

    def set_bottom_color(value)
      self.bottom_color = color(normalize_format_property_value(value))
    end

    def set_diag_color(value)
      self.diag_color = color(normalize_format_property_value(value))
    end

    def set_border_index(value)
      self.border_index = value
    end

    def set_border_count(value)
      self.border_count = value
    end

    #
    # Protection setters
    #
    def set_locked(value = 1)
      self.locked = value
    end

    def set_hidden(value = 1)
      self.hidden = value
    end

    #
    # Number format setters
    #
    def set_num_format(format)
      self.num_format = format
    end

    def set_num_format_index(value)
      self.num_format_index = value
    end

    #
    # Hyperlink and misc setters
    #
    def set_hyperlink(value)
      self.xf_id = 1

      set_underline(1)
      set_theme(10)
      self.hyperlink = value
    end

    def set_quote_prefix(value = 1)
      self.quote_prefix = value
    end

    def set_has_font(value = 1)
      self.has_font = value
    end

    def set_xf_index(value)
      self.xf_index = value
    end

    def set_dxf_index(value)
      self.dxf_index = value
    end

    def set_xf_id(value)
      self.xf_id = value
    end

    ###########################################################################
    #
    # Workbook integration and key generation
    #
    ###########################################################################

    def get_format_key
      [
        get_font_key,
        get_border_key,
        get_fill_key,
        get_alignment_key,
        num_format,
        locked,
        hidden,
        quote_prefix
      ].join(':')
    end

    def get_font_key
      [
        font,
        size,
        bold,
        italic,
        font_color,
        underline,
        font_strikeout,
        font_script,
        font_outline,
        font_shadow,
        font_family,
        font_charset,
        font_scheme,
        font_condense,
        font_extend,
        theme,
        hyperlink
      ].join(':')
    end

    def get_border_key
      [
        bottom,
        bottom_color,
        diag_border,
        diag_color,
        diag_type,
        left,
        left_color,
        right,
        right_color,
        top,
        top_color
      ].join(':')
    end

    def get_fill_key
      [
        pattern,
        bg_color,
        fg_color
      ].join(':')
    end

    def get_alignment_key
      [
        text_h_align,
        text_v_align,
        indent,
        rotation,
        text_wrap,
        shrink,
        reading_order
      ].join(':')
    end

    def get_xf_index
      if xf_index
        xf_index
      elsif @formats.xf_index_by_key(get_format_key)
        @formats.xf_index_by_key(get_format_key)
      else
        self.xf_index = @formats.set_xf_index_by_key(get_format_key)
      end
    end

    def get_dxf_index
      if dxf_index
        dxf_index
      elsif @formats.dxf_index_by_key(get_format_key)
        @formats.dxf_index_by_key(get_format_key)
      else
        self.dxf_index = @formats.set_dxf_index_by_key(get_format_key)
      end
    end

    def set_font_info(fonts)
      key = get_font_key

      if fonts[key]
        self.font_index = fonts[key]
        @has_font = false
      else
        self.font_index = fonts.size
        fonts[key] = fonts.size
        @has_font = true
      end
    end

    def set_border_info(borders)
      key = get_border_key

      if borders[key]
        self.border_index = borders[key]
        @has_border = false
      else
        self.border_index = borders.size
        borders[key] = borders.size
        @has_border = true
      end
    end

    ###########################################################################
    #
    # Queries and compatibility helpers
    #
    ###########################################################################

    def inspect
      to_s
    end

    def color(color_code)
      Format.color(color_code)
    end

    def color?
      ptrue?(font_color)
    end

    def bold?
      ptrue?(bold)
    end

    def italic?
      ptrue?(italic)
    end

    def strikeout?
      ptrue?(font_strikeout)
    end

    def outline?
      ptrue?(font_outline)
    end

    def shadow?
      ptrue?(font_shadow)
    end

    def underline?
      ptrue?(underline)
    end

    def has_border(flag)
      @has_border = flag
    end

    def has_border?
      @has_border
    end

    def has_dxf_border(flag)
      @has_dxf_border = flag
    end

    def has_dxf_border?
      @has_dxf_border
    end

    def has_font=(flag)
      @has_font = flag
    end

    def has_font?
      @has_font
    end

    def has_dxf_font(flag)
      @has_dxf_font = flag
    end

    def has_dxf_font?
      @has_dxf_font
    end

    def has_fill(flag)
      @has_fill = flag
    end

    def has_fill?
      @has_fill
    end

    def has_dxf_fill(flag)
      @has_dxf_fill = flag
    end

    def has_dxf_fill?
      @has_dxf_fill
    end

    #
    # Compatibility accessor used by older internal callers that treat Format
    # like a hash.
    #
    def [](attr)
      case attr.to_sym
      when :font           then font
      when :size           then size
      when :bold           then bold
      when :italic         then italic
      when :color          then font_color
      when :underline      then underline
      when :font_strikeout then font_strikeout
      when :font_script    then font_script
      when :rotation       then rotation
      else
        instance_variable_get("@#{attr}")
      end
    end

    def force_text_format?
      num_format == 49
    end

    ###########################################################################
    #
    # Alignment and protection serialization helpers
    #
    ###########################################################################

    def get_align_properties
      align = []

      h_align            = text_h_align
      v_align            = text_v_align
      indent_value       = indent
      rotation_value     = rotation
      wrap               = text_wrap
      shrink_value       = shrink
      reading_value      = reading_order
      just_distrib_value = just_distrib

      if h_align != 0 ||
         v_align != 0 ||
         indent_value != 0 ||
         rotation_value != 0 ||
         wrap != 0 ||
         shrink_value != 0 ||
         reading_value != 0
        changed = 1
      else
        return
      end

      if indent_value != 0 && ![1, 3, 7].include?(h_align) && ![1, 3, 5].include?(v_align)
        h_align = 1
      end

      shrink_value       = 0 if wrap != 0
      shrink_value       = 0 if h_align == 4
      shrink_value       = 0 if h_align == 5
      shrink_value       = 0 if h_align == 7
      just_distrib_value = 0 if h_align != 7
      just_distrib_value = 0 if indent_value != 0

      continuous = 'centerContinuous'

      align << %w[horizontal left]        if h_align == 1
      align << %w[horizontal center]      if h_align == 2
      align << %w[horizontal right]       if h_align == 3
      align << %w[horizontal fill]        if h_align == 4
      align << %w[horizontal justify]     if h_align == 5
      align << ['horizontal', continuous] if h_align == 6
      align << %w[horizontal distributed] if h_align == 7

      align << ['justifyLastLine', 1] if just_distrib_value != 0

      align << %w[vertical top]         if v_align == 1
      align << %w[vertical center]      if v_align == 2
      align << %w[vertical justify]     if v_align == 4
      align << %w[vertical distributed] if v_align == 5

      align << ['textRotation', rotation_value] if rotation_value != 0
      align << ['indent', indent_value]         if indent_value != 0

      align << ['wrapText', 1]    if wrap != 0
      align << ['shrinkToFit', 1] if shrink_value != 0

      align << ['readingOrder', 1] if reading_value == 1
      align << ['readingOrder', 2] if reading_value == 2

      [changed, align]
    end

    def get_protection_properties
      return if locked != 0 && hidden == 0

      attributes = []
      attributes << ['locked', 0] if locked == 0
      attributes << ['hidden', 1] if hidden != 0
      attributes
    end

    ###########################################################################
    #
    # XML writing entry points
    #
    ###########################################################################

    def write_font(writer, worksheet, dxf_format = nil) # :nodoc:
      writer.tag_elements('font') do
        write_condense(writer) if ptrue?(font_condense)
        write_extend(writer)   if ptrue?(font_extend)

        write_font_shapes(writer)

        writer.empty_tag('sz', [['val', size]]) unless dxf_format

        if theme == -1
          # Ignore for excel2003_style
        elsif ptrue?(theme)
          write_color('theme', theme, writer)
        elsif ptrue?(color_indexed)
          write_color('indexed', color_indexed, writer)
        elsif ptrue?(font_color)
          color = worksheet.palette_color(font_color)
          write_color('rgb', color, writer) if color != 'Automatic'
        elsif !ptrue?(dxf_format)
          write_color('theme', 1, writer)
        end

        unless ptrue?(dxf_format)
          writer.empty_tag('name', [['val', font]])
          write_font_family_scheme(writer)
        end
      end
    end

    def write_font_rpr(writer, worksheet) # :nodoc:
      writer.tag_elements('rPr') do
        write_font_shapes(writer)
        writer.empty_tag('sz', [['val', size]])

        if ptrue?(theme)
          write_color('theme', theme, writer)
        elsif ptrue?(font_color)
          color = worksheet.palette_color(font_color)
          write_color('rgb', color, writer)
        else
          write_color('theme', 1, writer)
        end

        writer.empty_tag('rFont', [['val', font]])
        write_font_family_scheme(writer)
      end
    end

    def border_attributes
      attributes = []

      if diag_type == 1
        attributes << ['diagonalUp', 1]
      elsif diag_type == 2
        attributes << ['diagonalDown', 1]
      elsif diag_type == 3
        attributes << ['diagonalUp', 1]
        attributes << ['diagonalDown', 1]
      end

      attributes
    end

    def xf_attributes
      attributes = [
        ['numFmtId', num_format_index],
        ['fontId', font_index],
        ['fillId', fill_index],
        ['borderId', border_index],
        ['xfId', xf_id]
      ]

      attributes << ['quotePrefix', 1] if ptrue?(quote_prefix)
      attributes << ['applyNumberFormat', 1] if num_format_index > 0
      attributes << ['applyFont', 1] if font_index > 0 && !ptrue?(hyperlink)
      attributes << ['applyFill', 1] if fill_index > 0
      attributes << ['applyBorder', 1] if border_index > 0

      apply_align, _align = get_align_properties
      attributes << ['applyAlignment', 1] if apply_align || ptrue?(hyperlink)
      attributes << ['applyProtection', 1] if get_protection_properties || ptrue?(hyperlink)

      attributes
    end

    ###########################################################################
    #
    # Class-level utilities
    #
    ###########################################################################

    def self.color(color_code)
      colors = Colors::COLORS

      return 0x00 unless color_code

      if color_code.respond_to?(:to_str)
        return color_code if color_code =~ /^#[0-9A-F]{6}$/i
        return colors[color_code.downcase.to_sym] if colors[color_code.downcase.to_sym]

        0x00 if color_code =~ /\D/
      else
        return color_code + 8 if color_code < 8
        return 0x00 if color_code > 63

        color_code
      end
    end

    ###########################################################################
    #
    # Private helpers
    #
    ###########################################################################
    private

    def normalize_format_property_value(value)
      if value.respond_to?(:to_str) || !value.respond_to?(:+)
        value.to_s
      else
        value
      end
    end

    ###########################################################################
    #
    # Private XML font helpers
    #
    ###########################################################################

    def write_font_shapes(writer)
      writer.empty_tag('b')       if bold?
      writer.empty_tag('i')       if italic?
      writer.empty_tag('strike')  if strikeout?
      writer.empty_tag('outline') if outline?
      writer.empty_tag('shadow')  if shadow?

      write_underline(writer, underline) if underline?

      write_vert_align(writer, 'superscript') if font_script == 1
      write_vert_align(writer, 'subscript')   if font_script == 2
    end

    def write_font_family_scheme(writer)
      writer.empty_tag('family', [['val', font_family]]) if ptrue?(font_family)
      writer.empty_tag('charset', [['val', font_charset]]) if ptrue?(font_charset)
      writer.empty_tag('scheme', [['val', font_scheme]]) if font == 'Calibri' && !ptrue?(hyperlink)
    end

    def write_underline(writer, underline)
      writer.empty_tag('u', write_underline_attributes(underline))
    end

    def write_underline_attributes(underline)
      val = 'val'

      case underline
      when 2
        [[val, 'double']]
      when 33
        [[val, 'singleAccounting']]
      when 34
        [[val, 'doubleAccounting']]
      else
        []
      end
    end

    def write_vert_align(writer, val) # :nodoc:
      writer.empty_tag('vertAlign', [['val', val]])
    end

    def write_condense(writer)
      writer.empty_tag('condense', [['val', 0]])
    end

    def write_extend(writer)
      writer.empty_tag('extend', [['val', 0]])
    end
  end
end
