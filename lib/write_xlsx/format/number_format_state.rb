# -*- coding: utf-8 -*-
# frozen_string_literal: true

module Writexlsx
  class Format
    class NumberFormatState
      attr_accessor :format_code, :index

      def initialize
        @format_code = 'General'
        @index       = 0
      end

      def initialize_copy(other)
        @format_code = other.format_code
        @index       = other.index
      end
    end
  end
end
