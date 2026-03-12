# -*- coding: utf-8 -*-
# frozen_string_literal: true

module Writexlsx
  class Workbook
    module Initialization
      #
      # user must not use. it is internal method.
      #
      def assemble_xml_file  # :nodoc:
        return unless @writer

        # Prepare format object for passing to Style.rb.
        prepare_format_properties

        write_xml_declaration do
          # Write the root workbook element.
          write_workbook do
            # Write the XLSX file version.
            write_file_version

            # Write the fileSharing element.
            write_file_sharing

            # Write the workbook properties.
            write_workbook_pr

            # Write the workbook view properties.
            write_book_views

            # Write the worksheet names and ids.
            @worksheets.write_sheets(@writer)

            # Write the workbook defined names.
            write_defined_names

            # Write the workbook calculation properties.
            write_calc_pr

            # Write the workbook extension storage.
            # write_ext_lst
          end
        end
      end

      private

      def write_workbook(&block) # :nodoc:
        schema = 'http://schemas.openxmlformats.org'
        attributes = [
          ['xmlns',
           schema + '/spreadsheetml/2006/main'],
          ['xmlns:r',
           schema + '/officeDocument/2006/relationships']
        ]
        @writer.tag_elements('workbook', attributes, &block)
      end

      def write_file_version # :nodoc:
        attributes = [
          %w[appName xl],
          ['lastEdited', 4],
          ['lowestEdited', 4],
          ['rupBuild', 4505]
        ]

        attributes << [:codeName, '{37E998C4-C9E5-D4B9-71C8-EB1FF731991C}'] if @vba_project

        @writer.empty_tag('fileVersion', attributes)
      end

      #
      # Write the <fileSharing> element.
      #
      def write_file_sharing
        return unless ptrue?(@read_only)

        attributes = []
        attributes << ['readOnlyRecommended', 1]
        @writer.empty_tag('fileSharing', attributes)
      end

      def write_workbook_pr # :nodoc:
        attributes = []
        attributes << ['codeName', @vba_codename]  if ptrue?(@vba_codename)
        attributes << ['date1904', 1]              if date_1904?
        attributes << ['defaultThemeVersion', 124226]
        @writer.empty_tag('workbookPr', attributes)
      end

      def write_book_views # :nodoc:
        @writer.tag_elements('bookViews') { write_workbook_view }
      end

      def write_workbook_view # :nodoc:
        attributes = [
          ['xWindow',       @x_window],
          ['yWindow',       @y_window],
          ['windowWidth',   @window_width],
          ['windowHeight',  @window_height]
        ]
        attributes << ['tabRatio', @tab_ratio] if @tab_ratio != 600
        attributes << ['firstSheet', @firstsheet + 1] if @firstsheet > 0
        attributes << ['activeTab', @activesheet] if @activesheet > 0
        @writer.empty_tag('workbookView', attributes)
      end

      def write_calc_pr # :nodoc:
        attributes = [['calcId', @calc_id]]

        case @calc_mode
        when 'manual'
          attributes << %w[calcMode manual]
          attributes << ['calcOnSave', 0]
        when 'autoNoTable'
          attributes << %w[calcMode autoNoTable]
        end

        attributes << ['fullCalcOnLoad', 1] if @calc_on_load

        @writer.empty_tag('calcPr', attributes)
      end

      def write_ext_lst # :nodoc:
        @writer.tag_elements('extLst') { write_ext }
      end

      def write_ext # :nodoc:
        attributes = [
          ['xmlns:mx', "#{OFFICE_URL}mac/excel/2008/main"],
          ['uri', uri]
        ]
        @writer.tag_elements('ext', attributes) { write_mx_arch_id }
      end

      def write_mx_arch_id # :nodoc:
        @writer.empty_tag('mx:ArchID', ['Flags', 2])
      end

      def write_defined_names # :nodoc:
        return unless ptrue?(@defined_names)

        @writer.tag_elements('definedNames') do
          @defined_names.each { |defined_name| write_defined_name(defined_name) }
        end
      end

      def write_defined_name(defined_name) # :nodoc:
        name, id, range, hidden = defined_name

        attributes = [['name', name]]
        attributes << ['localSheetId', id.to_s] unless id == -1
        attributes << %w[hidden 1]     if hidden

        @writer.data_element('definedName', range, attributes)
      end

      def write_io(str) # :nodoc:
        @writer << str
        str
      end
    end
  end
end
