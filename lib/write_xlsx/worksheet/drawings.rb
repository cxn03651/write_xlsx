# frozen_string_literal: true

module Writexlsx
  class Worksheet
    # Drawing insertion operations extracted from Worksheet to slim the main class.
    module DrawingMethods
      #
      # This method can be used to insert a Chart object into a worksheet.
      # The Chart must be created by the add_chart() Workbook method and
      # it must have the embedded option set.
      #
      def insert_chart(row, col, chart = nil, *options)
        # Check for a cell reference in A1 notation and substitute row and column.
        if (row_col_array = row_col_notation(row))
          _row, _col = row_col_array
          _chart     = col
          _options   = [chart] + options
        else
          _row = row
          _col = col
          _chart = chart
          _options = options
        end
        raise WriteXLSXInsufficientArgumentError if [_row, _col, _chart].include?(nil)

        if _options.first.instance_of?(Hash)
          params = _options.first
          x_offset    = params[:x_offset]
          y_offset    = params[:y_offset]
          x_scale     = params[:x_scale]
          y_scale     = params[:y_scale]
          anchor      = params[:object_position]
          description = params[:description]
          decorative  = params[:decorative]
        else
          x_offset, y_offset, x_scale, y_scale, anchor = _options
        end
        x_offset ||= 0
        y_offset ||= 0
        x_scale  ||= 1
        y_scale  ||= 1
        anchor   ||= 1

        raise "Not a Chart object in insert_chart()" unless _chart.is_a?(Chart) || _chart.is_a?(Chartsheet)
        raise "Not a embedded style Chart object in insert_chart()" if _chart.respond_to?(:embedded) && _chart.embedded == 0

        if _chart.already_inserted? || (_chart.combined && _chart.combined.already_inserted?)
          raise "Chart cannot be inserted in a worksheet more than once"
        else
          _chart.already_inserted          = true
          _chart.combined.already_inserted = true if _chart.combined
        end

        # Use the values set with chart.set_size, if any.
        x_scale  = _chart.x_scale  if _chart.x_scale  != 1
        y_scale  = _chart.y_scale  if _chart.y_scale  != 1
        x_offset = _chart.x_offset if ptrue?(_chart.x_offset)
        y_offset = _chart.y_offset if ptrue?(_chart.y_offset)

        @charts << InsertedChart.new(
          _row,    _col,    _chart, x_offset,    y_offset,
          x_scale, y_scale, anchor, description, decorative
        )
      end

      #
      # :call-seq:
      #   insert_image(row, column, filename, options)
      #
      def insert_image(row, col, image = nil, *options)
        # Check for a cell reference in A1 notation and substitute row and column.
        if (row_col_array = row_col_notation(row))
          _row, _col = row_col_array
          _image     = col
          _options   = [image] + options
        else
          _row = row
          _col = col
          _image = image
          _options = options
        end
        raise WriteXLSXInsufficientArgumentError if [_row, _col, _image].include?(nil)

        if _options.first.instance_of?(Hash)
          # Newer hash bashed options
          params      = _options.first
          x_offset    = params[:x_offset]
          y_offset    = params[:y_offset]
          x_scale     = params[:x_scale]
          y_scale     = params[:y_scale]
          anchor      = params[:object_position]
          url         = params[:url]
          tip         = params[:tip]
          description = params[:description]
          decorative  = params[:decorative]
        else
          x_offset, y_offset, x_scale, y_scale, anchor = _options
        end
        x_offset ||= 0
        y_offset ||= 0
        x_scale  ||= 1
        y_scale  ||= 1
        anchor   ||= 2

        @images << Image.new(
          _row, _col, _image, x_offset, y_offset,
          x_scale, y_scale, url, tip, anchor, description, decorative
        )
      end

      #
      # Embed an image into the worksheet.
      #
      def embed_image(row, col, filename, options = nil)
        # Check for a cell reference in A1 notation and substitute row and column.
        if (row_col_array = row_col_notation(row))
          _row, _col = row_col_array
          image      = col
          _options   = filename
        else
          _row     = row
          _col     = col
          image    = filename
          _options = options
        end
        xf, url, tip, description, decorative = []

        raise WriteXLSXInsufficientArgumentError if [_row, _col, image].include?(nil)
        raise "Couldn't locate #{image}" unless File.exist?(image)

        # Check that row and col are valid and store max and min values
        check_dimensions(_row, _col)
        store_row_col_max_min_values(_row, _col)

        if options
          xf          = options[:cell_format]
          url         = options[:url]
          tip         = options[:tip]
          description = options[:description]
          decorative  = options[:decorative]
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
        store_data_to_table(EmbedImageCellData.new(image_index, xf), _row, _col)
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
        # Check for a cell reference in A1 notation and substitute row and column.
        if (row_col_array = row_col_notation(row_start))
          _row_start, _column_start = row_col_array
          _shape    = column_start
          _x_offset = shape
          _y_offset = x_offset
          _x_scale  = y_offset
          _y_scale  = x_scale
          _anchor   = y_scale
        else
          _row_start = row_start
          _column_start = column_start
          _shape = shape
          _x_offset = x_offset
          _y_offset = y_offset
          _x_scale = x_scale
          _y_scale = y_scale
          _anchor = anchor
        end
        raise "Insufficient arguments in insert_shape()" if [_row_start, _column_start, _shape].include?(nil)

        _shape.set_position(
          _row_start, _column_start, _x_offset, _y_offset,
          _x_scale, _y_scale, _anchor
        )
        # Assign a shape ID.
        loop do
          id = _shape.id || 0
          used = @shape_hash[id]

          # Test if shape ID is already used. Otherwise assign a new one.
          if !used && id != 0
            break
          else
            @last_shape_id += 1
            _shape.id = @last_shape_id
          end
        end

        # Allow lookup of entry into shape array by shape ID.
        @shape_hash[_shape.id] = _shape.element = @shapes.size

        insert = if ptrue?(_shape.stencil)
                   # Insert a copy of the shape, not a reference so that the shape is
                   # used as a stencil. Previously stamped copies don't get modified
                   # if the stencil is modified.
                   _shape.dup
                 else
                   _shape
                 end

        # For connectors change x/y coords based on location of connected shapes.
        insert.auto_locate_connectors(@shapes, @shape_hash)

        # Insert a link to the shape on the list of shapes. Connection to
        # the parent shape is maintained.
        @shapes << insert
        insert
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
        @tables << table
        table
      end

      #
      # :call-seq:
      #    add_sparkline(properties)
      #
      # Add sparklines to the worksheet.
      #
      def add_sparkline(param)
        @sparklines << Sparkline.new(self, param, quote_sheetname(@name))
      end
    end
  end
end
