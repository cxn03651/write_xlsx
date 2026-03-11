# -*- coding: utf-8 -*-
# frozen_string_literal: true

module Writexlsx
  class Worksheet
    module DrawingPreparation
      ###############################################################################
      #
      # DrawingPreparation
      #
      # Prepares drawing-related assets for XLSX output.
      #
      # Responsibilities:
      # - Convert inserted charts, images, shapes, and media into Drawing objects
      # - Register drawing assets with the Workbook and Drawings containers
      # - Prepare background, header/footer images, and VML objects
      # - Coordinate drawing setup before XML serialization
      #
      # This module handles *how drawing assets are assembled for output*.
      # Relationship management and XML writing are handled elsewhere.
      #
      ###############################################################################

      # Drawing preparation

      def prepare_drawings(drawing_id, chart_ref_id, image_ref_id, image_ids, header_image_ids, background_ids)
        has_drawings = false

        # Check that some image or drawing needs to be processed.
        unless some_image_or_drawing_to_be_processed?

          # Don't increase the drawing_id header/footer images.
          unless charts.empty? && images.empty? && shapes.empty?
            drawing_id += 1
            has_drawings = true
          end

          # Prepare the background images.
          image_ref_id = prepare_background_image(background_ids, image_ref_id)

          # Prepare the worksheet images.
          images.each do |image|
            image_ref_id = prepare_image(image, drawing_id, image_ids, image_ref_id)
          end

          # Prepare the worksheet charts.
          charts.each_with_index do |_chart, index|
            chart_ref_id += 1
            prepare_chart(index, chart_ref_id, drawing_id)
          end

          # Prepare the worksheet shapes.
          shapes.each_with_index do |_shape, index|
            prepare_shape(index, drawing_id)
          end

          # Prepare the header and footer images.
          [header_images, footer_images].each do |images|
            images.each do |image|
              image_ref_id = prepare_header_footer_image(
                image, header_image_ids, image_ref_id
              )
            end
          end

          if has_drawings
            @workbook.drawings << drawings
          end
        end

        [drawing_id, chart_ref_id, image_ref_id]
      end

      #
      # Set up chart/drawings.
      #
      def prepare_chart(index, chart_id, drawing_id) # :nodoc:
        drawing_type = 1

        inserted_chart = charts[index]
        inserted_chart.chart.id = chart_id - 1

        dimensions = position_object_emus(inserted_chart)

        # Create a Drawing object to use with worksheet unless one already exists.
        drawing = Drawing.new(
          drawing_type, dimensions, 0, 0, nil, inserted_chart.anchor,
          drawing_rel_index, 0, nil, inserted_chart.name,
          inserted_chart.description, inserted_chart.decorative
        )
        if drawings?
          @drawings.add_drawing_object(drawing)
        else
          @drawings = Drawings.new
          @drawings.add_drawing_object(drawing)
          @drawings.embedded = true

          @external_drawing_links << ['/drawing', "../drawings/drawing#{drawing_id}.xml"]
        end
        @drawing_links << ['/chart', "../charts/chart#{chart_id}.xml"]
      end

      #
      # Set up image/drawings.
      #
      def prepare_image(image, drawing_id, image_ids, image_ref_id) # :nodoc:
        image_type = image.type
        x_dpi  = image.x_dpi || 96
        y_dpi  = image.y_dpi || 96
        md5    = image.md5
        drawing_type = 2

        @workbook.store_image_types(image_type)

        if image_ids[md5]
          image_id = image_ids[md5]
        else
          image_ref_id += 1
          image_ids[md5] = image_id = image_ref_id
          @workbook.images << image
        end

        dimensions = position_object_emus(image)

        # Create a Drawing object to use with worksheet unless one already exists.
        drawing = Drawing.new(
          drawing_type, dimensions, image.width_emus, image.height_emus,
          nil, image.anchor, 0, 0, image.tip, image.name,
          image.description || image.name, image.decorative
        )
        unless drawings?
          @drawings = Drawings.new
          @drawings.embedded = true

          @external_drawing_links << ['/drawing', "../drawings/drawing#{drawing_id}.xml"]
        end
        @drawings.add_drawing_object(drawing)

        if image.url
          target_mode = 'External'
          target = escape_url(image.url) if image.url =~ %r{^[fh]tt?ps?://} || image.url =~ /^mailto:/
          if image.url =~ /^external:/
            target = escape_url(image.url.sub(/^external:/, ''))

            # Additional escape not required in worksheet hyperlinks
            target = target.gsub("#", '%23')

            # Prefix absolute paths (not relative) with file:///
            target = if target =~ /^\w:/ || target =~ /^\\\\/
                       "file:///#{target}"
                     else
                       target.gsub("\\", '/')
                     end
          end

          if image.url =~ /^internal:/
            target      = image.url.sub(/^internal:/, '#')
            target_mode = nil
          end

          if target.length > 255
            raise <<"EOS"
Ignoring URL #{target} where link or anchor > 255 characters since it exceeds Excel's limit for URLS. See LIMITATIONS section of the WriteXLSX documentation.
EOS
          end

          @drawing_links << ['/hyperlink', target, target_mode] if target && !@drawing_rels[image.url]
          drawing.url_rel_index = drawing_rel_index(image.url)
        end

        @drawing_links << ['/image', "../media/image#{image_id}.#{image_type}"] unless @drawing_rels[md5]

        drawing.rel_index = drawing_rel_index(md5)

        image_ref_id
      end

      #
      # Set up drawing shapes
      #
      def prepare_shape(index, drawing_id)
        shape = shapes[index]

        # Create a Drawing object to use with worksheet unless one already exists.
        unless drawings?
          @drawings = Drawings.new
          @drawings.embedded = true
          @external_drawing_links << ['/drawing', "../drawings/drawing#{drawing_id}.xml"]
          @has_shapes = true
        end

        # Validate the he shape against various rules.
        shape.validate(index)
        shape.calc_position_emus(self)

        drawing_type = 3
        drawing = Drawing.new(
          drawing_type, shape.dimensions, shape.width_emu, shape.height_emu,
          shape, shape.anchor, drawing_rel_index, 0, shape.name, nil, 0
        )
        drawings.add_drawing_object(drawing)
      end

      # Image preparation

      #
      # Set up an image without a drawing object for the background image.
      #
      def prepare_background(image_id, image_type)
        @external_background_links <<
          ['/image', "../media/image#{image_id}.#{image_type}"]
      end

      def prepare_background_image(background_ids, image_ref_id)
        unless background_image.nil?
          @workbook.store_image_types(background_image.type)

          if background_ids[background_image.md5]
            ref_id = background_ids[background_image.md5]
          else
            image_ref_id += 1
            ref_id = image_ref_id
            background_ids[background_image.md5] = ref_id
            @workbook.images << background_image
          end

          prepare_background(ref_id, background_image.type)
        end

        image_ref_id
      end

      def prepare_header_image(image_id, image_property)
        # Strip the extension from the filename.
        body = image_property.name.dup
        body[/\.[^.]+$/, 0] = ''
        image_property.body = body

        @vml_drawing_links << ['/image', "../media/image#{image_id}.#{image_property.type}"] unless @vml_drawing_rels[image_property.md5]

        image_property.ref_id = get_vml_drawing_rel_index(image_property.md5)
        @header_images_array << image_property
      end

      def prepare_header_footer_image(image, header_image_ids, image_ref_id)
        @workbook.store_image_types(image.type)

        if header_image_ids[image.md5]
          ref_id = header_image_ids[image.md5]
        else
          image_ref_id += 1
          header_image_ids[image.md5] = ref_id = image_ref_id
          @workbook.images << image
        end

        prepare_header_image(ref_id, image)

        image_ref_id
      end

      # VML preparation

      #
      # Turn the HoH that stores the comments into an array for easier handling
      # and set the external links for comments and buttons.
      #
      def prepare_vml_objects(vml_data_id, vml_shape_id, vml_drawing_id, comment_id)
        set_external_vml_links(vml_drawing_id)
        set_external_comment_links(comment_id) if has_comments?

        # The VML o:idmap data id contains a comma separated range when there is
        # more than one 1024 block of comments, like this: data="1,2".
        data = vml_data_id.to_s
        (1..num_comments_block).each do |i|
          data += ",#{vml_data_id + i}"
        end
        @vml_data_id = data
        @vml_shape_id = vml_shape_id
      end

      #
      # Setup external linkage for VML header/footer images.
      #
      def prepare_header_vml_objects(vml_header_id, vml_drawing_id)
        @vml_header_id = vml_header_id
        @external_vml_links << ['/vmlDrawing', "../drawings/vmlDrawing#{vml_drawing_id}.vml"]
      end
    end
  end
end
