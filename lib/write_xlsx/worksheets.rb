# -*- encoding: utf-8 -*-
require 'delegate'
require 'write_xlsx/package/xml_writer_simple'

module Writexlsx
  class Worksheets < DelegateClass(Array)
    def initialize
      super([])
    end

    def chartsheet_count
      self.select { |worksheet| worksheet.is_chartsheet? }.count
    end

    def is_sheetname_uniq?(name)
      self.each do |worksheet|
        return false if name.downcase == worksheet.name.downcase
      end
      true
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

    private

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
