# -*- coding: utf-8 -*-
require 'write_xlsx/package/xml_writer_simple'
require 'write_xlsx/utility'
require 'write_xlsx/package/app'
require 'write_xlsx/package/comments'
require 'write_xlsx/package/content_types'
require 'write_xlsx/package/core'
require 'write_xlsx/package/relationships'
require 'write_xlsx/package/shared_strings'
require 'write_xlsx/package/styles'
require 'write_xlsx/package/theme'
require 'write_xlsx/package/vml'

module Writexlsx
  module Package
    class Packager

      include Writexlsx::Utility

      def initialize
        @package_dir       = ''
        @workbook         = nil
        @sheet_names      = []
        @worksheet_count  = 0
        @chartsheet_count = 0
        @chart_count      = 0
        @drawing_count    = 0
        @named_ranges     = []
      end

      def set_package_dir(package_dir)
        @package_dir = package_dir
      end

      #
      # Add the Workbook object to the package.
      #
      def add_workbook(workbook)
        @workbook          = workbook
        @sheet_names       = workbook.sheetnames
        @chart_count       = workbook.charts.size
        @drawing_count     = workbook.drawings.size
        @num_comment_files = workbook.num_comment_files
        @named_ranges      = workbook.named_ranges

        workbook.worksheets.each do |worksheet|
          if worksheet.is_chartsheet?
            @chartsheet_count += 1
          else
            @worksheet_count += 1
          end
        end
      end

      #
      # Write the xml files that make up the XLXS OPC package.
      #
      def create_package
        write_worksheet_files
        write_chartsheet_files
        write_workbook_file
        write_chart_files
        write_drawing_files
        write_vml_files
        write_comment_files
        write_shared_strings_file
        write_app_file
        write_core_file
        write_content_types_file
        write_styles_file
        write_theme_file
        write_root_rels_file
        write_workbook_rels_file
        write_worksheet_rels_files
        write_chartsheet_rels_files
        write_drawing_rels_files
        add_image_files
      end

      private

      #
      # Write the workbook.xml file.
      #
      def write_workbook_file
        FileUtils.mkdir_p("#{@package_dir}/xl")

        @workbook.set_xml_writer("#{@package_dir}/xl/workbook.xml")
        @workbook.assemble_xml_file
      end

      #
      # Write the worksheet files.
      #
      def write_worksheet_files
        FileUtils.mkdir_p("#{@package_dir}/xl/worksheets")

        index = 1
        @workbook.worksheets.each do |worksheet|
          next if worksheet.is_chartsheet?
          worksheet.set_xml_writer("#{@package_dir}/xl/worksheets/sheet#{index}.xml")
          index += 1
          worksheet.assemble_xml_file
        end
      end

      #
      def write_chartsheet_files
        index = 1
        @workbook.worksheets.each do |worksheet|
          next unless worksheet.is_chartsheet?
          FileUtils.mkdir_p("#{@package_dir}/xl/chartsheets")
          worksheet.set_xml_writer("#{@package_dir}/xl/chartsheets/sheet#{index}.xml")
          index += 1
          worksheet.assemble_xml_file
        end
      end

      #
      # Write the chart files.
      #
      def write_chart_files
        write_chart_or_drawing_files(@workbook.charts, 'chart')
      end

      #
      # Write the drawing files.
      #
      def write_drawing_files
        write_chart_or_drawing_files(@workbook.drawings, 'drawing')
      end

      def write_chart_or_drawing_files(objects, filename)
        return if objects.empty?

        FileUtils.mkdir_p("#{@package_dir}/xl/#{filename}s")

        index = 1
        objects.each do |object|
          object.set_xml_writer("#{@package_dir}/xl/#{filename}s/#{filename}#{index}.xml")
          index += 1
          object.assemble_xml_file
        end
      end

      #
      # Write the comment VML files.
      #
      def write_vml_files
        index = 1
        @workbook.worksheets.each do |worksheet|
          next unless worksheet.has_comments?
          FileUtils.mkdir_p("#{@package_dir}/xl/drawings")

          vml = Package::Vml.new
          vml.set_xml_writer("#{@package_dir}/xl/drawings/vmlDrawing#{index}.vml")
          index += 1
          vml.assemble_xml_file(worksheet)
        end
      end

      #
      # Write the comment files.
      #
      def write_comment_files
        index = 1
        @workbook.worksheets.each do |worksheet|
          next unless worksheet.has_comments?

          FileUtils.mkdir_p("#{@package_dir}/xl/drawings")

          worksheet.comments_xml_writer = "#{@package_dir}/xl/comments#{index}.xml"
          index += 1

          worksheet.comments_assemble_xml_file
        end
      end

      #
      # Write the sharedStrings.xml file.
      #
      def write_shared_strings_file
        sst  = @workbook.shared_strings

        FileUtils.mkdir_p("#{@package_dir}/xl")

        return if @workbook.shared_strings_empty?

        sst.set_xml_writer("#{@package_dir}/xl/sharedStrings.xml")
        sst.assemble_xml_file
      end

      #
      # Write the app.xml file.
      #
      def write_app_file
        dir        = @package_dir
        properties = @workbook.doc_properties
        app        = Package::App.new

        FileUtils.mkdir_p("#{@package_dir}/docProps")

        # Add the Worksheet heading pairs.
        app.add_heading_pair(['Worksheets', @worksheet_count])

        # Add the Chartsheet heading pairs.
        app.add_heading_pair(['Charts', @chartsheet_count])

        # Add the Worksheet parts.
        @workbook.worksheets.each do |worksheet|
          next if worksheet.is_chartsheet?
          app.add_part_name(worksheet.name)
        end

        # Add the Chartsheet parts.
        @workbook.worksheets.each do |worksheet|
          next unless worksheet.is_chartsheet?
          app.add_part_name(worksheet.get_name)
        end

        # Add the Named Range heading pairs.
        range_count = @named_ranges.size
        if range_count != 0
          app.add_heading_pair([ 'Named Ranges', range_count ])
        end

        # Add the Named Ranges parts.
        @named_ranges.each { |named_range| app.add_part_name(named_range) }

        app.set_properties(properties)

        app.set_xml_writer("#{@package_dir}/docProps/app.xml")
        app.assemble_xml_file
      end

      #
      # Write the core.xml file.
      #
      def write_core_file
        core       = Package::Core.new

        FileUtils.mkdir_p("#{@package_dir}/docProps")

        core.set_properties(@workbook.doc_properties)
        core.set_xml_writer("#{@package_dir}/docProps/core.xml")
        core.assemble_xml_file
      end

      #
      # Write the ContentTypes.xml file.
      #
      def write_content_types_file
        content = Package::ContentTypes.new

        content.add_image_types(@workbook.image_types)

        worksheet_index  = 1
        chartsheet_index = 1
        @workbook.worksheets.each do |worksheet|
          if worksheet.is_chartsheet?
            content.add_chartsheet_name("sheet#{chartsheet_index}")
            chartsheet_index += 1
          else
            content.add_worksheet_name("sheet#{worksheet_index}")
            worksheet_index += 1
          end
        end

        (1 .. @chart_count).each { |i| content.add_chart_name("chart#{i}") }
        (1 .. @drawing_count).each { |i| content.add_drawing_name("drawing#{i}") }

        content.add_vml_name if @num_comment_files > 0

        (1 .. @num_comment_files).each { |i| content.add_comment_name("comments#{i}") }

        # Add the sharedString rel if there is string data in the workbook.
        content.add_shared_strings unless @workbook.shared_strings_empty?

        content.set_xml_writer("#{@package_dir}/[Content_Types].xml")
        content.assemble_xml_file
      end

      #
      # Write the style xml file.
      #
      def write_styles_file
        dir              = @package_dir
        xf_formats       = @workbook.xf_formats
        palette          = @workbook.palette
        font_count       = @workbook.font_count
        num_format_count = @workbook.num_format_count
        border_count     = @workbook.border_count
        fill_count       = @workbook.fill_count
        custom_colors    = @workbook.custom_colors
        dxf_formats      = @workbook.dxf_formats

        rels = Package::Styles.new

        FileUtils.mkdir_p("#{@package_dir}/xl")

        rels.set_style_properties(
            xf_formats,
            palette,
            font_count,
            num_format_count,
            border_count,
            fill_count,
            custom_colors,
            dxf_formats
        )

        rels.set_xml_writer("#{@package_dir}/xl/styles.xml" )
        rels.assemble_xml_file
      end

      #
      # Write the style xml file.
      #
      def write_theme_file
        rels = Package::Theme.new

        FileUtils.mkdir_p("#{@package_dir}/xl/theme")

        rels.set_xml_writer("#{@package_dir}/xl/theme/theme1.xml")
        rels.assemble_xml_file
      end

      #
      # Write the _rels/.rels xml file.
      #
      def write_root_rels_file
        rels = Package::Relationships.new

        FileUtils.mkdir_p("#{@package_dir}/_rels")

        rels.add_document_relationship('/officeDocument', 'xl/workbook.xml')
        rels.add_package_relationship('/metadata/core-properties',
            'docProps/core')
        rels.add_document_relationship('/extended-properties', 'docProps/app.xml')
        rels.set_xml_writer("#{@package_dir}/_rels/.rels" )
        rels.assemble_xml_file
      end

      #
      # Write the _rels/.rels xml file.
      #
      def write_workbook_rels_file
        rels = Package::Relationships.new

        FileUtils.mkdir_p("#{@package_dir}/xl/_rels")

        worksheet_index  = 1
        chartsheet_index = 1

        @workbook.worksheets.each do |worksheet|
          if worksheet.is_chartsheet?
            rels.add_document_relationship('/chartsheet', "chartsheets/sheet#{chartsheet_index}.xml")
            chartsheet_index += 1
          else
            rels.add_document_relationship( '/worksheet', "worksheets/sheet#{worksheet_index}.xml")
            worksheet_index += 1
          end
        end

        rels.add_document_relationship('/theme',  'theme/theme1.xml')
        rels.add_document_relationship('/styles', 'styles.xml')

        # Add the sharedString rel if there is string data in the workbook.
        rels.add_document_relationship('/sharedStrings', 'sharedStrings.xml') unless @workbook.shared_strings_empty?
        rels.set_xml_writer("#{@package_dir}/xl/_rels/workbook.xml.rels")
        rels.assemble_xml_file
      end

      #
      # Write the worksheet .rels files for worksheets that contain links to external
      # data such as hyperlinks or drawings.
      #
      def write_worksheet_rels_files
        existing_rels_dir = false

        index = 0
        @workbook.worksheets.each do |worksheet|
          next if worksheet.is_chartsheet?

          index += 1

          external_links = [
            worksheet.external_hyper_links,
            worksheet.external_drawing_links,
            worksheet.external_comment_links
          ].select {|a| a != []}

          next if external_links.size == 0

          # Create the worksheet .rels dir if required.
          if !existing_rels_dir
            FileUtils.mkdir_p("#{@package_dir}/xl/worksheets")
            FileUtils.mkdir_p("#{@package_dir}/xl/worksheets/_rels")
            existing_rels_dir = true
          end

          rels = Package::Relationships.new

          external_links.each do |link_datas|
            link_datas.each do |link_data|
              type, target, target_mode = link_data
              rels.add_worksheet_relationship(type, target, target_mode)
            end
          end

          # Create the .rels file such as /xl/worksheets/_rels/sheet1.xml.rels.
          rels.set_xml_writer(
            "#{@package_dir}/xl/worksheets/_rels/sheet#{index}.xml.rels")
          rels.assemble_xml_file
        end
      end

      #
      # Write the chartsheet .rels files for links to drawing files.
      #
      def write_chartsheet_rels_files
        existing_rels_dir = false

        @workbook.worksheets.each do |worksheet|
          next unless worksheet.is_chartsheet?

          external_links = worksheet.external_drawing_links

          next if external_links.empty?

          # Create the chartsheet .rels dir if required.
          if existing_rels_dir
            FileUtils.mkdir_p("#{@package_dir}/xl/chartsheets/_rels")
            existing_rels_dir = true
          end

          rels = Package::Relationships.new

          external_links.each do |link_data|
            rels.add_worksheet_relationship(link_data)
          end

          # Create the .rels file such as /xl/chartsheets/_rels/sheet1.xml.rels.
          rels.set_xml_writer(
              "#{@package_dir}/xl/chartsheets/_rels/sheet#{worksheet.index}.xml.rels")
          rels.assemble_xml_file
        end
      end

      #
      # Write the drawing .rels files for worksheets that contain charts or drawings.
      #
      def write_drawing_rels_files
        index = 0
        @workbook.worksheets.each do |worksheet|
          next if worksheet.drawing_links.empty?
          index += 1

          # Create the drawing .rels dir if required.
          FileUtils.mkdir_p("#{@package_dir}/xl/drawings/_rels")

          rels = Package::Relationships.new

          worksheet.drawing_links.each do |drawing_data|
            rels.add_document_relationship(*drawing_data)
          end

          # Create the .rels file such as /xl/drawings/_rels/sheet1.xml.rels.
          rels.set_xml_writer(
            "#{@package_dir}/xl/drawings/_rels/drawing#{index}.xml.rels")
          rels.assemble_xml_file
        end
      end


      #
      # Write the workbook.xml file.
      #
      def add_image_files
        return if @workbook.images.empty?

        index    = 1

        FileUtils.mkdir_p("#{@package_dir}/xl/media")

        @workbook.images.each do |image|
          filename  = image[0]
          extension = ".#{image[1]}"

          copy( filename, "#{@package_dir}/xl/media/image#{index}#{extension}")
          index += 1
        end
      end
    end
  end
end
