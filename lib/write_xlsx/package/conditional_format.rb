# -*- coding: utf-8 -*-

module Writexlsx
  module Package
    class ConditionalFormat
      attr_reader :param

      def initialize(range, param)
        @range, @param = range, param
      end
    end
  end
end
