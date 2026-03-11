# -*- coding: utf-8 -*-
# frozen_string_literal: true

module Writexlsx
  class Worksheet
    module DrawingMethods
      ###############################################################################
      #
      # DrawingMethods
      #
      # Provides high-level Worksheet APIs for inserting drawing-related objects
      # such as charts, images, shapes, tables, and sparklines.
      #
      # Responsibilities:
      # - Public insertion entry points used by Worksheet users
      # - Argument normalization and option extraction
      # - Shape insertion workflow and helper utilities
      #
      # This module handles *what gets added to the sheet*.
      # It does not manage relationships or XML output.
      #
      ###############################################################################
      #
      # This method can be used to insert a Chart object into a worksheet.
      # The Chart must be created by the add_chart() Workbook method and
      # it must have the embedded option set.
      #
      def insert_chart(row, col, chart = nil, *options)
        normalized_row, normalized_col, normalized_chart, normalized_options =
          normalize_row_col_args(row, col, chart, options)
        raise WriteXLSXInsufficientArgumentError if [normalized_row, normalized_col, normalized_chart].include?(nil)

        x_offset, y_offset, x_scale, y_scale, anchor, description, decorative =
          extract_chart_options(normalized_options)

        raise "Not a Chart object in insert_chart()" unless normalized_chart.is_a?(Chart) || normalized_chart.is_a?(Chartsheet)
        raise "Not a embedded style Chart object in insert_chart()" if normalized_chart.respond_to?(:embedded) && normalized_chart.embedded == 0

        if normalized_chart.already_inserted? || (normalized_chart.combined && normalized_chart.combined.already_inserted?)
          raise "Chart cannot be inserted in a worksheet more than once"
        else
          normalized_chart.already_inserted          = true
          normalized_chart.combined.already_inserted = true if normalized_chart.combined
        end

        # Use the values set with chart.set_size, if any.
        x_scale  = normalized_chart.x_scale  if normalized_chart.x_scale  != 1
        y_scale  = normalized_chart.y_scale  if normalized_chart.y_scale  != 1
        x_offset = normalized_chart.x_offset if ptrue?(normalized_chart.x_offset)
        y_offset = normalized_chart.y_offset if ptrue?(normalized_chart.y_offset)

        @assets.add_chart(
          InsertedChart.new(
            normalized_row,    normalized_col, normalized_chart,
            x_offset, y_offset, x_scale, y_scale, anchor, description, decorative
          )
        )
      end

      def insert_image(row, col, image = nil, *options)
        normalized_row, normalized_col, normalized_image, normalized_options =
          normalize_row_col_args(row, col, image, options)
        raise WriteXLSXInsufficientArgumentError if [normalized_row, normalized_col, normalized_image].include?(nil)

        x_offset, y_offset, x_scale, y_scale,
        anchor, url, tip, description, decorative = extract_image_options(normalized_options)

        @assets.add_image(
          Image.new(
            normalized_row, normalized_col, normalized_image, x_offset, y_offset,
            x_scale, y_scale, url, tip, anchor, description, decorative
          )
        )
      end

      def embed_image(row, col, filename, options = nil)
        normalize_row, normalize_col, image, normalize_options = normalize_row_col_args(row, col, filename, options)

        raise WriteXLSXInsufficientArgumentError if [normalize_row, normalize_col, image].include?(nil)
        raise "Couldn't locate #{image}" unless File.exist?(image)

        # Check that row and col are valid and store max and min values
        check_dimensions(normalize_row, normalize_col)
        store_row_col_max_min_values(normalize_row, normalize_col)

        if options
          xf          = options[:cell_format]
          url         = options[:url]
          tip         = options[:tip]
          description = options[:description]
          decorative  = options[:decorative]
        else
          xf, url, tip, description, decorative = []
        end

        # Write the url without writing a string.
        if url
          xf ||= @default_url_format

          write_url(row, col, url, xf, nil, tip, true)
        end

        # Get the image properties, mainly for the type and checksum.
        image_property = ImageProperty.new(
          image, description: description, decorative: decorative
        )
        @workbook.store_image_types(image_property.type)

        # Check for duplicate images.
        image_index = @embedded_image_indexes[image_property.md5]

        unless ptrue?(image_index)
          @workbook.embedded_images << image_property

          image_index = @workbook.embedded_images.size
          @embedded_image_indexes[image_property.md5] = image_index
        end

        # Write the cell placeholder.
        store_data_to_table(EmbedImageCellData.new(image_index, xf), normalize_row, normalize_col)
        @has_embedded_images = true
      end

      #
      # :call-seq:
      #   insert_shape(row, col, shape [ , x, y, x_scale, y_scale ])
      #
      # Insert a shape into the worksheet.
      #
      def insert_shape(
            row_start, column_start, shape = nil, x_offset = nil, y_offset = nil,
            x_scale = nil, y_scale = nil, anchor = nil
          )
        row, col, normalized_shape, normalized_options =
          normalize_shape_args(row_start, column_start, shape, x_offset, y_offset, x_scale, y_scale, anchor)
        raise "Insufficient arguments in insert_shape()" if [row, col, normalized_shape].include?(nil)

        set_shape_position(normalized_shape, row, col, normalized_options)
        assign_shape_id(normalized_shape)

        # Allow lookup of entry into shape array by shape ID.
        @shape_hash[normalized_shape.id] = normalized_shape.element = shapes.size

        inserted_shape = build_inserted_shape(normalized_shape)

        # For connectors change x/y coords based on location of connected shapes.
        auto_locate_shape_connectors(inserted_shape)

        # Insert a link to the shape on the list of shapes. Connection to
        # the parent shape is maintained.
        @assets.add_shape(inserted_shape)
        inserted_shape
      end

      #
      # :call-seq:
      #    add_table(row1, col1, row2, col2, properties)
      #
      # Add an Excel table to a worksheet.
      #
      def add_table(*args)
        # Table count is a member of Workbook, global to all Worksheet.
        table = Package::Table.new(self, *args)
        @assets.add_table(table)
        table
      end

      #
      # :call-seq:
      #    add_sparkline(properties)
      #
      # Add sparklines to the worksheet.
      #
      def add_sparkline(param)
        @assets.add_sparkline(Sparkline.new(self, param, quote_sheetname(@name)))
      end

      private

      def normalize_row_col_args(row, col, object, options)
        if (row_col_array = row_col_notation(row))
          normalized_row, normalized_col = row_col_array
          normalized_object = col
          normalized_options = [object] + options
        else
          normalized_row = row
          normalized_col = col
          normalized_object = object
          normalized_options = options
        end

        [normalized_row, normalized_col, normalized_object, normalized_options]
      end

      def extract_chart_options(options)
        if options.first.instance_of?(Hash)
          params = options.first
          [
            params[:x_offset] || 0,
            params[:y_offset] || 0,
            params[:x_scale] || 1,
            params[:y_scale] || 1,
            params[:object_position] || 1,
            params[:description],
            params[:decorative]
          ]
        else
          x_offset, y_offset, x_scale, y_scale, anchor = options
          [x_offset || 0, y_offset || 0, x_scale || 1, y_scale || 1, anchor || 1, nil, nil]
        end
      end

      def extract_image_options(options)
        if options.first.instance_of?(Hash)
          params = options.first
          [
            params[:x_offset] || 0,
            params[:y_offset] || 0,
            params[:x_scale] || 1,
            params[:y_scale] || 1,
            params[:object_position] || 2,
            params[:url],
            params[:tip],
            params[:description],
            params[:decorative]
          ]
        else
          x_offset, y_offset, x_scale, y_scale, anchor = options
          [x_offset || 0, y_offset || 0, x_scale || 1, y_scale || 1, anchor || 2, nil, nil, nil, nil]
        end
      end

      def normalize_shape_args(row_start, column_start, shape,
                               x_offset, y_offset, x_scale, y_scale, anchor)
        if (row_col_array = row_col_notation(row_start))
          normalized_row, normalized_col = row_col_array
          normalized_shape    = column_start
          normalized_x_offset = shape
          normalized_y_offset = x_offset
          normalized_x_scale  = y_offset
          normalized_y_scale  = x_scale
          normalized_anchor   = y_scale
        else
          normalized_row      = row_start
          normalized_col      = column_start
          normalized_shape    = shape
          normalized_x_offset = x_offset
          normalized_y_offset = y_offset
          normalized_x_scale  = x_scale
          normalized_y_scale  = y_scale
          normalized_anchor   = anchor
        end

        [
          normalized_row,
          normalized_col,
          normalized_shape,
          {
            x_offset: normalized_x_offset,
            y_offset: normalized_y_offset,
            x_scale:  normalized_x_scale,
            y_scale:  normalized_y_scale,
            anchor:   normalized_anchor
          }
        ]
      end

      def set_shape_position(shape, row, col, options)
        shape.set_position(
          row, col,
          options[:x_offset],
          options[:y_offset],
          options[:x_scale],
          options[:y_scale],
          options[:anchor]
        )
      end

      def assign_shape_id(shape)
        loop do
          id = shape.id || 0
          used = @shape_hash[id]

          if !used && id != 0
            break
          else
            @last_shape_id += 1
            shape.id = @last_shape_id
          end
        end

        @shape_hash[shape.id] = shape.element = shapes.size
      end

      def build_inserted_shape(shape)
        if ptrue?(shape.stencil)
          shape.dup
        else
          shape
        end
      end

      def auto_locate_shape_connectors(shape)
        shape.auto_locate_connectors(shapes, @shape_hash)
      end
    end
  end
end
