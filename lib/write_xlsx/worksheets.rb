# -*- encoding: utf-8 -*-
require 'delegate'
require 'write_xlsx/package/xml_writer_simple'

module Writexlsx
  class Worksheets < DelegateClass(Array)
    include Writexlsx::Utility

    BASE_NAME = { :sheet => 'Sheet', :chart => 'Chart'}  # :nodoc:

    def initialize
      super([])
    end

    def chartsheet_count
      self.select { |worksheet| worksheet.is_chartsheet? }.count
    end

    def sheetname_count
      self.count - chartname_count
    end

    def chartname_count
      chartsheet_count
    end

    def make_and_check_sheet_chart_name(type, name)
      count = sheet_chart_count(type)
      name = "#{BASE_NAME[type]}#{count+1}" unless ptrue?(name)

      check_valid_sheetname(name)
      name
    end

    def write_sheets(writer)
      writer.tag_elements('sheets') do
        id_num = 1
        self.each do |sheet|
          write_sheet(writer, sheet, id_num)
          id_num += 1
        end
      end
    end

    def write_worksheet_files(package_dir)
      dir = "#{package_dir}/xl/worksheets"
      self.reject { |sheet| sheet.is_chartsheet? }.
        each_with_index do |sheet, index|
        write_sheet_files(dir, sheet, index)
      end
    end

    def write_chartsheet_files(package_dir)
      dir = "#{package_dir}/xl/chartsheets"
      self.select { |sheet| sheet.is_chartsheet? }.
        each_with_index do |sheet, index|
        write_sheet_files(dir, sheet, index)
      end
    end

    def write_vml_files(package_dir)
      dir = "#{package_dir}/xl/drawings"
      self.select { |sheet| sheet.has_vml? }.
        each_with_index do |sheet, index|
        FileUtils.mkdir_p(dir)

        vml = Package::Vml.new
        vml.set_xml_writer("#{dir}/vmlDrawing#{index+1}.vml")
        vml.assemble_xml_file(sheet)
      end
    end

    def write_comment_files(package_dir)
      self.select { |sheet| sheet.has_comments? }.
        each_with_index do |sheet, index|
        FileUtils.mkdir_p("#{package_dir}/xl/drawings")
        sheet.comments_xml_writer = "#{package_dir}/xl/comments#{index+1}.xml"
        sheet.comments_assemble_xml_file
      end
    end

    def write_table_files(package_dir)
      unless tables.empty?
        dir = "#{package_dir}/xl/tables"
        FileUtils.mkdir_p(dir)
        tables.each_with_index do |table, index|
          table.set_xml_writer("#{dir}/table#{index+1}.xml")
          table.assemble_xml_file
        end
      end
    end

    def write_chartsheet_rels_files(package_dir)
      dir = "#{package_dir}/xl/chartsheets/_rels"
      self.select { |sheet| sheet.is_chartsheet? }.
        each_with_index do |sheet, index|

        external_links = sheet.external_drawing_links

        next if external_links.empty?

        FileUtils.mkdir_p(dir)
        rels = Package::Relationships.new

        external_links.each do |link_data|
          rels.add_worksheet_relationship(*link_data)
        end

        # Create the .rels file such as /xl/chartsheets/_rels/sheet1.xml.rels.
        rels.set_xml_writer("#{dir}/sheet#{index+1}.xml.rels")
        rels.assemble_xml_file
      end
    end

    def write_drawing_rels_files(package_dir)
      dir = "#{package_dir}/xl/drawings/_rels"
      self.reject { |sheet| sheet.drawing_links.empty? }.
        each_with_index do |sheet, index|

        FileUtils.mkdir_p(dir)

        rels = Package::Relationships.new

        sheet.drawing_links.each do |drawing_data|
          rels.add_document_relationship(*drawing_data)
        end

        # Create the .rels file such as /xl/drawings/_rels/sheet1.xml.rels.
        rels.set_xml_writer("#{dir}/drawing#{index+1}.xml.rels")
        rels.assemble_xml_file
      end
    end

    def write_worksheet_rels_files(package_dir)
      dir = "#{package_dir}/xl/worksheets/_rels"

      self.reject { |sheet| sheet.is_chartsheet? }.
        each_with_index do |sheet, index|

        external_links = [
                          sheet.external_hyper_links,
                          sheet.external_drawing_links,
                          sheet.external_vml_links,
                          sheet.external_table_links,
                          sheet.external_comment_links
                         ].reject { |a| a.empty? }

        next if external_links.size == 0

        FileUtils.mkdir_p(dir)

        rels = Package::Relationships.new

        external_links.each do |link_datas|
          link_datas.each do |link_data|
            rels.add_worksheet_relationship(*link_data)
          end
        end

        # Create the .rels file such as /xl/worksheets/_rels/sheet1.xml.rels.
        rels.set_xml_writer("#{dir}/sheet#{index+1}.xml.rels")
        rels.assemble_xml_file
      end
    end

    def tables
      self.inject([]) { |tables, sheet| tables + sheet.tables }.flatten
    end

    def tables_count
      tables.count
    end

    def index_by_name(sheetname)
      name = sheetname.sub(/^'/,'').sub(/'$/,'')
      self.collect { |sheet| sheet.name }.index(name)
    end

    private

    def sheet_chart_count(type)
      case type
      when :sheet
        sheetname_count
      when :chart
        chartname_count
      end
    end

    def check_valid_sheetname(name)
      # Check that sheet name is <= 31. Excel limit.
      raise "Sheetname #{name} must be <= #{SHEETNAME_MAX} chars" if name.length > SHEETNAME_MAX

      # Check that sheetname doesn't contain any invalid characters
      invalid_char = /[\[\]:*?\/\\]/
      if name =~ invalid_char
        raise 'Invalid character []:*?/\\ in worksheet name: ' + name
      end

      # Check that the worksheet name doesn't already exist since this is a fatal
      # error in Excel 97. The check must also exclude case insensitive matches.
      unless is_sheetname_uniq?(name)
        raise "Worksheet name '#{name}', with case ignored, is already used."
      end
    end

    def is_sheetname_uniq?(name)
      self.each do |worksheet|
        return false if name.downcase == worksheet.name.downcase
      end
      true
    end

    def write_sheet_files(dir, sheet, index)
      FileUtils.mkdir_p(dir)
      sheet.set_xml_writer("#{dir}/sheet#{index+1}.xml")
      sheet.assemble_xml_file
    end

    def write_sheet(writer, sheet, sheet_id) #:nodoc:
      attributes = [
        'name',    sheet.name,
        'sheetId', sheet_id
      ]

      if sheet.hidden?
        attributes << 'state' << 'hidden'
      end
      attributes << 'r:id' << "rId#{sheet_id}"
      writer.empty_tag_encoded('sheet', attributes)
    end
  end
end
