# -*- coding: utf-8 -*-
# frozen_string_literal: true

module Writexlsx
  class Format
    class NumberFormatStyle
      def initialize(format)
        @format = format
      end

      def format_code
        @format.instance_variable_get(:@number_format_state).format_code
      end

      def format_code=(value)
        @format.instance_variable_get(:@number_format_state).format_code = value
        @format.send(:sync_number_format_ivars_from_state)
      end

      def index
        @format.instance_variable_get(:@number_format_state).index
      end

      def index=(value)
        @format.instance_variable_get(:@number_format_state).index = value
        @format.send(:sync_number_format_ivars_from_state)
      end
    end
  end
end
