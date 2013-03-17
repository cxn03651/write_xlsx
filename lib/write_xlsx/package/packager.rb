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
require 'write_xlsx/package/table'
require 'write_xlsx/package/theme'
require 'write_xlsx/package/vml'

module Writexlsx
  module Package
    class Packager

      include Writexlsx::Utility

      def initialize(workbook)
        @workbook     = workbook
        @package_dir  = ''
        @table_count  = @workbook.worksheets.tables_count
        @named_ranges = []
      end

      def set_package_dir(package_dir)
        @package_dir = package_dir
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
        write_table_files
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
        add_vba_project
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
        @workbook.worksheets.write_worksheet_files(@package_dir)
      end

      #
      def write_chartsheet_files
        @workbook.worksheets.write_chartsheet_files(@package_dir)
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
        dir = "#{@package_dir}/xl/#{filename}s"

        objects.each_with_index do |object, index|
          FileUtils.mkdir_p(dir)
          object.set_xml_writer("#{dir}/#{filename}#{index+1}.xml")
          object.assemble_xml_file
        end
      end

      #
      # Write the comment VML files.
      #
      def write_vml_files
        @workbook.worksheets.write_vml_files(@package_dir)
      end

      #
      # Write the comment files.
      #
      def write_comment_files
        @workbook.worksheets.write_comment_files(@package_dir)
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
        app.add_heading_pair(['Worksheets', @workbook.worksheets.reject {|s| s.is_chartsheet?}.count])

        # Add the Chartsheet heading pairs.
        app.add_heading_pair(['Charts', @workbook.chartsheet_count])

        # Add the Worksheet parts.
        @workbook.worksheets.reject { |sheet| sheet.is_chartsheet? }.
          each { |sheet| app.add_part_name(sheet.name) }

        # Add the Chartsheet parts.
        @workbook.worksheets.select { |sheet| sheet.is_chartsheet? }.
          each { |sheet| app.add_part_name(sheet.name) }

        # Add the Named Range heading pairs.
        range_count = @workbook.named_ranges.size
        if range_count != 0
          app.add_heading_pair([ 'Named Ranges', range_count ])
        end

        # Add the Named Ranges parts.
        @workbook.named_ranges.each { |named_range| app.add_part_name(named_range) }

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
        @workbook.worksheets.reject { |sheet| sheet.is_chartsheet? }.
          each_with_index do |sheet, index|
          content.add_worksheet_name("sheet#{index+1}")
        end
        @workbook.worksheets.select { |sheet| sheet.is_chartsheet? }.
          each_with_index do |sheet, index|
          content.add_chartsheet_name("sheet#{index+1}")
        end

        (1 .. @workbook.charts.size).each { |i| content.add_chart_name("chart#{i}") }
        (1 .. @workbook.drawings.size).each { |i| content.add_drawing_name("drawing#{i}") }

        content.add_vml_name if @workbook.num_vml_files > 0

        (1 .. @table_count).each { |i| content.add_table_name("table#{i}") }

        (1 .. @workbook.num_comment_files).each { |i| content.add_comment_name("comments#{i}") }

        # Add the sharedString rel if there is string data in the workbook.
        content.add_shared_strings unless @workbook.shared_strings_empty?

        # Add vbaProject if present.
        content.add_vba_project if @workbook.vba_project

        content.set_xml_writer("#{@package_dir}/[Content_Types].xml")
        content.assemble_xml_file
      end

      #
      # Write the style xml file.
      #
      def write_styles_file
        dir              = "#{@package_dir}/xl"

        rels = Package::Styles.new

        FileUtils.mkdir_p(dir)

        rels.set_style_properties(*@workbook.style_properties)

        rels.set_xml_writer("#{dir}/styles.xml" )
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
      # Write the table files.
      #
      def write_table_files
        @workbook.worksheets.write_table_files(@package_dir)
      end

      #
      # Write the _rels/.rels xml file.
      #
      def write_root_rels_file
        rels = Package::Relationships.new

        FileUtils.mkdir_p("#{@package_dir}/_rels")

        rels.add_document_relationship('/officeDocument', 'xl/workbook.xml')
        rels.add_package_relationship('/metadata/core-properties',
            'docProps/core.xml')
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

        # Add vbaProject if present.
        if @workbook.vba_project
          rels.add_ms_package_relationship('/vbaProject', 'vbaProject.bin')
        end

        rels.set_xml_writer("#{@package_dir}/xl/_rels/workbook.xml.rels")
        rels.assemble_xml_file
      end

      #
      # Write the worksheet .rels files for worksheets that contain links to external
      # data such as hyperlinks or drawings.
      #
      def write_worksheet_rels_files
        @workbook.worksheets.write_worksheet_rels_files(@package_dir)
      end

      #
      # Write the chartsheet .rels files for links to drawing files.
      #
      def write_chartsheet_rels_files
        @workbook.worksheets.write_chartsheet_rels_files(@package_dir)
      end

      #
      # Write the drawing .rels files for worksheets that contain charts or drawings.
      #
      def write_drawing_rels_files
        @workbook.worksheets.write_drawing_rels_files(@package_dir)
      end


      #
      # Write the /xl/media/image?.xml files.
      #
      def add_image_files
        dir = "#{@package_dir}/xl/media"

        @workbook.images.each_with_index do |image, index|
          FileUtils.mkdir_p(dir)
          FileUtils.cp(image[0], "#{dir}/image#{index+1}.#{image[1]}")
        end
      end

      #
      # Write the vbaProject.bin file.
      #
      def add_vba_project
        dir = @package_dir
        vba_project = @workbook.vba_project

        return unless vba_project

        FileUtils.mkdir_p("#{dir}/xl")
        FileUtils.copy(vba_project, "#{dir}/xl/vbaProject.bin")
      end
    end
  end
end
