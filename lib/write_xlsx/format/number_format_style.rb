# -*- coding: utf-8 -*-
# frozen_string_literal: true

module Writexlsx
  class Format
    class NumberFormatStyle
      def initialize(format)
        @format = format
      end

      def format_code
        @format.state.number_format.format_code
      end

      def format_code=(value)
        @format.state.number_format.format_code = value
        @format.send(:sync_number_format_ivars_from_state)
      end

      def index
        @format.state.number_format.index
      end

      def index=(value)
        @format.state.number_format.index = value
        @format.send(:sync_number_format_ivars_from_state)
      end
    end
  end
end
