# -*- coding: utf-8 -*-
# frozen_string_literal: true

module Writexlsx
  class Chart
    class Caption
      include Writexlsx::Utility::Common
      include Writexlsx::Utility::RichText

      attr_accessor :name, :formula, :data_id, :font
      attr_accessor :line, :fill, :pattern, :gradient
      attr_reader :layout, :overlay, :none

      def initialize(chart)
        @chart = chart
      end

      def apply_options(params) # :nodoc:
        @name, @formula = chart.process_names(params[:name], params[:name_formula])
        @name = nil if @name.respond_to?(:empty?) && @name.empty?
        @data_id = chart.data_id(@formula, params[:data])
        @font     = convert_font_args(params[:font] || params[:name_font])

        @layout   = chart.layout_properties(params[:layout], 1)
        @overlay  = params[:overlay]
        @none     = params[:none]
      end

      def apply_format_options(params)
        @line     = chart.line_properties(params[:border] || params[:line])
        @fill     = chart.fill_properties(params[:fill])
        @pattern  = chart.pattern_properties(params[:pattern])
        @gradient = chart.gradient_properties(params[:gradient])
      end

      private

      attr_reader :chart
    end
  end
end
