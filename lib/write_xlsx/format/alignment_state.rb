# -*- coding: utf-8 -*-
# frozen_string_literal: true

module Writexlsx
  class Format
    class AlignmentState
      attr_accessor :horizontal, :vertical,
                    :wrap, :justlast, :rotation,
                    :indent, :shrink, :merge_range,
                    :reading_order, :just_distrib

      def initialize
        @horizontal    = 0
        @vertical      = 0
        @wrap          = 0
        @justlast      = 0
        @rotation      = 0
        @indent        = 0
        @shrink        = 0
        @merge_range   = 0
        @reading_order = 0
        @just_distrib  = 0
      end

      def initialize_copy(other)
        @horizontal    = other.horizontal
        @vertical      = other.vertical
        @wrap          = other.wrap
        @justlast      = other.justlast
        @rotation      = other.rotation
        @indent        = other.indent
        @shrink        = other.shrink
        @merge_range   = other.merge_range
        @reading_order = other.reading_order
        @just_distrib  = other.just_distrib
      end
    end
  end
end
