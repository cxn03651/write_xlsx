# -*- coding: utf-8 -*-

module Writexlsx
  class Chart
    class Caption
      attr_accessor :name, :formula, :data_id, :name_font
      attr_reader :layout, :overlay

      def initialize(chart)
        @chart = chart
      end

      def merge_with_hash(params) # :nodoc:
        @name, @formula = @chart.process_names(params[:name], params[:name_formula])
        @data_id        = @chart.get_data_id(@formula, params[:data])
        @name_font      = @chart.convert_font_args(params[:name_font])
        @layout   = @chart.layout_properties(params[:layout], 1)
        @overlay  = params[:overlay]
      end
    end
  end
end
