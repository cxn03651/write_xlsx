# -*- coding: utf-8 -*-
# frozen_string_literal: true

module Writexlsx
  class Format
    class FontState
      attr_accessor(
        :index,
        :name,
        :size,
        :bold,
        :italic,
        :color,
        :underline,
        :strikeout,
        :outline,
        :shadow,
        :script,
        :family,
        :charset,
        :scheme,
        :condense,
        :extend,
        :theme,
        :hyperlink,
        :color_indexed
      )

      def initialize
        @index         = 0
        @name          = 'Calibri'
        @size          = 11
        @bold          = 0
        @italic        = 0
        @color         = 0x0
        @underline     = 0
        @strikeout     = 0
        @outline       = 0
        @shadow        = 0
        @script        = 0
        @family        = 2
        @charset       = 0
        @scheme        = 'minor'
        @condense      = 0
        @extend        = 0
        @theme         = 0
        @hyperlink     = 0
        @color_indexed = 0
      end

      def initialize_copy(other)
        @index         = other.index
        @name          = other.name
        @size          = other.size
        @bold          = other.bold
        @italic        = other.italic
        @color         = other.color
        @underline     = other.underline
        @strikeout     = other.strikeout
        @outline       = other.outline
        @shadow        = other.shadow
        @script        = other.script
        @family        = other.family
        @charset       = other.charset
        @scheme        = other.scheme
        @condense      = other.condense
        @extend        = other.extend
        @theme         = other.theme
        @hyperlink     = other.hyperlink
        @color_indexed = other.color_indexed
      end
    end
  end
end
