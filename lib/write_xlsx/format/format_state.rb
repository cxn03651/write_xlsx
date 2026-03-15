# -*- coding: utf-8 -*-
# frozen_string_literal: true

module Writexlsx
  class Format
    class FormatState
      attr_accessor :fill, :border, :font, :alignment, :protection, :number_format
      attr_accessor :xf_index, :dxf_index, :xf_id
      attr_accessor :quote_prefix
      attr_accessor :has_fill, :has_font, :has_border
      attr_accessor :has_dxf_fill, :has_dxf_font, :has_dxf_border
      attr_accessor :dxf_fg_color, :dxf_bg_color

      def initialize
        @fill          = nil
        @border        = nil
        @font          = nil
        @alignment     = nil
        @protection    = nil
        @number_format = nil

        @xf_index       = nil
        @dxf_index      = nil
        @xf_id          = 0
        @quote_prefix   = 0
        @has_fill       = false
        @has_font       = false
        @has_border     = false
        @has_dxf_fill   = false
        @has_dxf_font   = false
        @has_dxf_border = false
        @dxf_fg_color   = nil
        @dxf_bg_color   = nil
      end

      def initialize_copy(other)
        @fill          = other.fill&.dup
        @border        = other.border&.dup
        @font          = other.font&.dup
        @alignment     = other.alignment&.dup
        @protection    = other.protection&.dup
        @number_format = other.number_format&.dup

        @xf_index       = other.xf_index
        @dxf_index      = other.dxf_index
        @xf_id          = other.xf_id
        @quote_prefix   = other.quote_prefix
        @has_fill       = other.has_fill
        @has_font       = other.has_font
        @has_border     = other.has_border
        @has_dxf_fill   = other.has_dxf_fill
        @has_dxf_font   = other.has_dxf_font
        @has_dxf_border = other.has_dxf_border
        @dxf_fg_color   = other.dxf_fg_color
        @dxf_bg_color   = other.dxf_bg_color
      end
    end
  end
end
