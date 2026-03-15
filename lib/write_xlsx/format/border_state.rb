# -*- coding: utf-8 -*-
# frozen_string_literal: true

module Writexlsx
  class Format
    class BorderState
      attr_accessor :index, :count,
                    :bottom, :bottom_color,
                    :diag_border, :diag_color, :diag_type,
                    :left, :left_color,
                    :right, :right_color,
                    :top, :top_color

      def initialize
        @index        = 0
        @count        = 0
        @bottom       = 0
        @bottom_color = 0x0
        @diag_border  = 0
        @diag_color   = 0x0
        @diag_type    = 0
        @left         = 0
        @left_color   = 0x0
        @right        = 0
        @right_color  = 0x0
        @top          = 0
        @top_color    = 0x0
      end

      def initialize_copy(other)
        @index        = other.index
        @count        = other.count
        @bottom       = other.bottom
        @bottom_color = other.bottom_color
        @diag_border  = other.diag_border
        @diag_color   = other.diag_color
        @diag_type    = other.diag_type
        @left         = other.left
        @left_color   = other.left_color
        @right        = other.right
        @right_color  = other.right_color
        @top          = other.top
        @top_color    = other.top_color
      end
    end
  end
end
