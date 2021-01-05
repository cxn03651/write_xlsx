# -*- coding: utf-8 -*-

module Writexlsx
  class Chart
    class Legend
      attr_accessor :line, :fill, :pattern, :gradient
      attr_accessor :position, :delete_series, :layout, :font

      def initialize
        @position = 'right'
      end
    end
  end
end
