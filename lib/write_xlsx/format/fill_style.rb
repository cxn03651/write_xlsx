# -*- coding: utf-8 -*-
# frozen_string_literal: true

module Writexlsx
  class Format
    class FillStyle
      def initialize(format)
        @format = format
      end

      def fg_color
        @format.instance_variable_get(:@fill).fg_color
      end

      def bg_color
        @format.instance_variable_get(:@fill).bg_color
      end

      def pattern
        @format.instance_variable_get(:@fill).pattern
      end

      def index
        @format.instance_variable_get(:@fill).index
      end

      def count
        @format.instance_variable_get(:@fill).count
      end

      def fg_color=(value)
        @format.instance_variable_get(:@fill).fg_color = value
        @format.send(:sync_fill_ivars_from_state)
      end

      def bg_color=(value)
        @format.instance_variable_get(:@fill).bg_color = value
        @format.send(:sync_fill_ivars_from_state)
      end

      def pattern=(value)
        @format.instance_variable_get(:@fill).pattern = value
        @format.send(:sync_fill_ivars_from_state)
      end

      def index=(value)
        @format.instance_variable_get(:@fill).index = value
        @format.send(:sync_fill_ivars_from_state)
      end

      def count=(value)
        @format.instance_variable_get(:@fill).count = value
        @format.send(:sync_fill_ivars_from_state)
      end
    end
  end
end
