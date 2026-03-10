# -*- coding: utf-8 -*-
# frozen_string_literal: true

module Writexlsx
  class Worksheet
    class AssetManager
      attr_reader :charts, :images, :tables, :sparklines, :shapes
      attr_accessor :background_image
      attr_reader :header_images, :footer_images

      def initialize
        @charts = []
        @images = []
        @tables = []
        @sparklines = []
        @shapes = []

        @header_images = []
        @footer_images = []
        @background_image = nil
      end

      def add_chart(chart)
        @charts << chart
      end

      def add_image(image)
        @images << image
      end

      def add_table(table)
        @tables << table
      end

      def add_sparkline(sparkline)
        @sparklines << sparkline
      end

      def add_shape(shape)
        @shapes << shape
      end

      def reset_header_images
        @header_images
      end

      def reset_footer_images
        @footer_images
      end

      def add_header_image(image)
        @header_images << image
      end

      def add_footer_image(image)
        @footer_images << image
      end
    end
  end
end
