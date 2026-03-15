# -*- coding: utf-8 -*-
# frozen_string_literal: true

module Writexlsx
  class Format
    class BorderStyle
      def initialize(format)
        @format = format
      end

      def index
        @format.instance_variable_get(:@border).index
      end

      def index=(value)
        @format.instance_variable_get(:@border).index = value
        @format.send(:sync_border_ivars_from_state)
      end

      def count
        @format.instance_variable_get(:@border).count
      end

      def count=(value)
        @format.instance_variable_get(:@border).count = value
        @format.send(:sync_border_ivars_from_state)
      end

      def left
        @format.instance_variable_get(:@border).left
      end

      def left=(value)
        @format.instance_variable_get(:@border).left = value
        @format.send(:sync_border_ivars_from_state)
      end

      def left_color
        @format.instance_variable_get(:@border).left_color
      end

      def left_color=(value)
        @format.instance_variable_get(:@border).left_color = value
        @format.send(:sync_border_ivars_from_state)
      end

      def right
        @format.instance_variable_get(:@border).right
      end

      def right=(value)
        @format.instance_variable_get(:@border).right = value
        @format.send(:sync_border_ivars_from_state)
      end

      def right_color
        @format.instance_variable_get(:@border).right_color
      end

      def right_color=(value)
        @format.instance_variable_get(:@border).right_color = value
        @format.send(:sync_border_ivars_from_state)
      end

      def top
        @format.instance_variable_get(:@border).top
      end

      def top=(value)
        @format.instance_variable_get(:@border).top = value
        @format.send(:sync_border_ivars_from_state)
      end

      def top_color
        @format.instance_variable_get(:@border).top_color
      end

      def top_color=(value)
        @format.instance_variable_get(:@border).top_color = value
        @format.send(:sync_border_ivars_from_state)
      end

      def bottom
        @format.instance_variable_get(:@border).bottom
      end

      def bottom=(value)
        @format.instance_variable_get(:@border).bottom = value
        @format.send(:sync_border_ivars_from_state)
      end

      def bottom_color
        @format.instance_variable_get(:@border).bottom_color
      end

      def bottom_color=(value)
        @format.instance_variable_get(:@border).bottom_color = value
        @format.send(:sync_border_ivars_from_state)
      end

      def diag_border
        @format.instance_variable_get(:@border).diag_border
      end

      def diag_border=(value)
        @format.instance_variable_get(:@border).diag_border = value
        @format.send(:sync_border_ivars_from_state)
      end

      def diag_color
        @format.instance_variable_get(:@border).diag_color
      end

      def diag_color=(value)
        @format.instance_variable_get(:@border).diag_color = value
        @format.send(:sync_border_ivars_from_state)
      end

      def diag_type
        @format.instance_variable_get(:@border).diag_type
      end

      def diag_type=(value)
        @format.instance_variable_get(:@border).diag_type = value
        @format.send(:sync_border_ivars_from_state)
      end
    end
  end
end
