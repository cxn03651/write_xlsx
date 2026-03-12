# -*- coding: utf-8 -*-
# frozen_string_literal: true

module Writexlsx
  class Chart
    class Table
      include Writexlsx::Utility

      attr_reader :horizontal, :vertical, :outline, :show_keys, :font

      def initialize(params = {})
        @horizontal = true
        @vertical   = true
        @outline    = true
        @show_keys  = false
        @horizontal = params[:horizontal] if params.has_key?(:horizontal)
        @vertical   = params[:vertical]   if params.has_key?(:vertical)
        @outline    = params[:outline]    if params.has_key?(:outline)
        @show_keys  = params[:show_keys]  if params.has_key?(:show_keys)
        @font       = convert_font_args(params[:font])
      end

      attr_writer :palette

      def write_d_table(writer)
        @writer = writer
        @writer.tag_elements('c:dTable') do
          @writer.empty_tag('c:showHorzBorder', attributes) if ptrue?(horizontal)
          @writer.empty_tag('c:showVertBorder', attributes) if ptrue?(vertical)
          @writer.empty_tag('c:showOutline',    attributes) if ptrue?(outline)
          @writer.empty_tag('c:showKeys',       attributes) if ptrue?(show_keys)
          # Write the table font.
          write_tx_pr(font)                                 if ptrue?(font)
        end
      end

      private

      def attributes
        [['val', 1]]
      end
    end
  end
end
