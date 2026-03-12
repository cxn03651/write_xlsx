# -*- coding: utf-8 -*-
# frozen_string_literal: true

module Writexlsx
  class Workbook
    module PackagePreparation
      private

      #
      # Assemble worksheets into a workbook.
      #
      def store_workbook # :nodoc:
        # Add a default worksheet if non have been added.
        add_worksheet if @worksheets.empty?

        # Ensure that at least one worksheet has been selected.
        @worksheets.visible_first.select if @activesheet == 0

        # Set the active sheet.
        @activesheet = @worksheets.visible_first.index if @activesheet == 0
        @worksheets[@activesheet].activate

        # Convert the SST strings data structure.
        prepare_sst_string_data

        # Prepare the worksheet VML elements such as comments and buttons.
        prepare_vml_objects
        # Set the defined names for the worksheets such as Print Titles.
        prepare_defined_names
        # Prepare the drawings, charts and images.
        prepare_drawings
        # Add cached data to charts.
        add_chart_data

        # Prepare the worksheet tables.
        prepare_tables

        # Prepare the metadata file links.
        prepare_metadata

        # Package the workbook.
        packager = Package::Packager.new(self)
        packager.set_package_dir(tempdir)
        packager.create_package

        # Free up the Packager object.
        packager = nil

        # Store the xlsx component files with the temp dir name removed.
        ZipFileUtils.zip(tempdir.to_s, filename)

        IO.copy_stream(filename, fileobj) if fileobj
        delete_tempdir(tempdir)
      end

      #
      # Iterate through the worksheets and store any defined names in addition to
      # any user defined names. Stores the defined names for the Workbook.xml and
      # the named ranges for App.xml.
      #
      def prepare_defined_names # :nodoc:
        @worksheets.each do |sheet|
          # Check for Print Area settings.
          if sheet.autofilter_area
            @defined_names << [
              '_xlnm._FilterDatabase',
              sheet.index,
              sheet.autofilter_area,
              1
            ]
          end

          # Check for Print Area settings.
          unless sheet.print_area.empty?
            @defined_names << [
              '_xlnm.Print_Area',
              sheet.index,
              sheet.print_area
            ]
          end

          # Check for repeat rows/cols. aka, Print Titles.
          next unless !sheet.print_repeat_cols.empty? || !sheet.print_repeat_rows.empty?

          range = if !sheet.print_repeat_cols.empty? && !sheet.print_repeat_rows.empty?
                    sheet.print_repeat_cols + ',' + sheet.print_repeat_rows
                  else
                    sheet.print_repeat_cols + sheet.print_repeat_rows
                  end

          # Store the defined names.
          @defined_names << ['_xlnm.Print_Titles', sheet.index, range]
        end

        @defined_names = sort_defined_names(@defined_names)
        @named_ranges  = extract_named_ranges(@defined_names)
      end

      #
      # Iterate through the worksheets and set up the VML objects.
      #
      def prepare_vml_objects  # :nodoc:
        comment_id     = 0
        vml_drawing_id = 0
        vml_data_id    = 1
        vml_header_id  = 0
        vml_shape_id   = 1024
        has_button     = false

        @worksheets.each do |sheet|
          next if !sheet.has_vml? && !sheet.has_header_vml?

          if sheet.has_vml?
            if sheet.has_comments?
              comment_id += 1
              @has_comments = true
            end
            vml_drawing_id += 1

            sheet.prepare_vml_objects(
              vml_data_id, vml_shape_id,
              vml_drawing_id, comment_id
            )

            # Each VML file should start with a shape id incremented by 1024.
            vml_data_id += 1 * (1 + sheet.num_comments_block)
            vml_shape_id += 1024 * (1 + sheet.num_comments_block)
          end

          if sheet.has_header_vml?
            vml_header_id  += 1
            vml_drawing_id += 1
            sheet.prepare_header_vml_objects(vml_header_id, vml_drawing_id)
          end

          # Set the sheet vba_codename if it has a button and the workbook
          # has a vbaProject binary.
          unless sheet.buttons_data.empty?
            has_button = true
            sheet.set_vba_name if @vba_project && !sheet.vba_codename
          end
        end

        # Set the workbook vba_codename if one of the sheets has a button and
        # the workbook has a vbaProject binary.
        set_vba_name if has_button && @vba_project && !@vba_codename
      end

      #
      # Set the table ids for the worksheet tables.
      #
      def prepare_tables
        table_id = 0
        seen     = {}

        sheets.each do |sheet|
          table_id += sheet.prepare_tables(table_id + 1, seen)
        end
      end

      #
      # Set the metadata rel link.
      #
      def prepare_metadata
        @worksheets.each do |sheet|
          next unless sheet.has_dynamic_functions? || sheet.has_embedded_images?

          @has_metadata = true
          @has_dynamic_functions ||= sheet.has_dynamic_functions?
          @has_embedded_images   ||= sheet.has_embedded_images?
        end
      end

      #
      # Iterate through the worksheets and set up any chart or image drawings.
      #
      def prepare_drawings # :nodoc:
        # Store the image types for any embedded images.
        @embedded_images.each do |image|
          store_image_types(image.type)

          @has_embedded_descriptions = true if ptrue?(image.description)
        end

        prepare_drawings_of_all_sheets

        # Sort the workbook charts references into the order that the were
        # written from the worksheets above.
        @charts = @charts.reject { |chart| chart.id == -1 }
                    .sort_by(&:id)
      end

      def prepare_drawings_of_all_sheets
        drawing_id       = 0
        chart_ref_id     = 0
        image_ids        = {}
        header_image_ids = {}
        background_ids   = {}

        # The image IDs start from after the embedded images.
        image_ref_id = @embedded_images.size

        @worksheets.each do |sheet|
          drawing_id, chart_ref_id, image_ref_id =
            sheet.prepare_drawings(
              drawing_id, chart_ref_id, image_ref_id, image_ids,
              header_image_ids, background_ids
            )
        end
      end

      #
      # prepare_sst_string_data
      #
      def prepare_sst_string_data; end

      def delete_tempdir(path)
        if FileTest.file?(path)
          File.delete(path)
        elsif FileTest.directory?(path)
          Dir.foreach(path) do |file|
            next if file =~ /^\.\.?$/  # '.' or '..'

            delete_tempdir(path.sub(%r{/+$}, "") + '/' + file)
          end
          Dir.rmdir(path)
        end
      end
    end
  end
end
