# -*- coding: utf-8 -*-
# frozen_string_literal: true

module Writexlsx
  class Format
    class FillStyle
      def initialize(format)
        @format = format
      end

      def fg_color
        @format.state.fill.fg_color
      end

      def bg_color
        @format.state.fill.bg_color
      end

      def pattern
        @format.state.fill.pattern
      end

      def index
        @format.state.fill.index
      end

      def count
        @format.state.fill.count
      end

      def fg_color=(value)
        @format.state.fill.fg_color = value
      end

      def bg_color=(value)
        @format.state.fill.bg_color = value
      end

      def pattern=(value)
        @format.state.fill.pattern = value
      end

      def index=(value)
        @format.state.fill.index = value
      end

      def count=(value)
        @format.state.fill.count = value
      end
    end
  end
end
