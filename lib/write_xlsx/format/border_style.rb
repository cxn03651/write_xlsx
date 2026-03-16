# -*- coding: utf-8 -*-
# frozen_string_literal: true

module Writexlsx
  class Format
    class BorderStyle
      def initialize(format)
        @format = format
      end

      def index
        @format.state.border.index
      end

      def index=(value)
        @format.state.border.index = value
      end

      def count
        @format.state.border.count
      end

      def count=(value)
        @format.state.border.count = value
      end

      def left
        @format.state.border.left
      end

      def left=(value)
        @format.state.border.left = value
      end

      def left_color
        @format.state.border.left_color
      end

      def left_color=(value)
        @format.state.border.left_color = value
      end

      def right
        @format.state.border.right
      end

      def right=(value)
        @format.state.border.right = value
      end

      def right_color
        @format.state.border.right_color
      end

      def right_color=(value)
        @format.state.border.right_color = value
      end

      def top
        @format.state.border.top
      end

      def top=(value)
        @format.state.border.top = value
      end

      def top_color
        @format.state.border.top_color
      end

      def top_color=(value)
        @format.state.border.top_color = value
      end

      def bottom
        @format.state.border.bottom
      end

      def bottom=(value)
        @format.state.border.bottom = value
      end

      def bottom_color
        @format.state.border.bottom_color
      end

      def bottom_color=(value)
        @format.state.border.bottom_color = value
      end

      def diag_border
        @format.state.border.diag_border
      end

      def diag_border=(value)
        @format.state.border.diag_border = value
      end

      def diag_color
        @format.state.border.diag_color
      end

      def diag_color=(value)
        @format.state.border.diag_color = value
      end

      def diag_type
        @format.state.border.diag_type
      end

      def diag_type=(value)
        @format.state.border.diag_type = value
      end
    end
  end
end
