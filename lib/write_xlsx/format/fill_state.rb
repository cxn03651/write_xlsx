# -*- coding: utf-8 -*-
# frozen_string_literal: true

module Writexlsx
  class Format
    class FillState
      attr_accessor :fg_color, :bg_color, :pattern, :index, :count

      def initialize
        @fg_color = 0x00
        @bg_color = 0x00
        @pattern  = 0
        @index    = 0
        @count    = 0
      end

      def initialize_copy(other)
        @fg_color = other.fg_color
        @bg_color = other.bg_color
        @pattern  = other.pattern
        @index    = other.index
        @count    = other.count
      end
    end
  end
end
