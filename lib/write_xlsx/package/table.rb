# -*- coding: utf-8 -*-

require 'write_xlsx/package/xml_writer_simple'
require 'write_xlsx/utility'

module Writexlsx
  module Package
    class Table
      attr_writer :properties

      def initialize
        @writer = Package::XMLWriterSimple.new
        @properties = {}
      end

      #
      # Assemble and writes the XML file.
      #
      def assemble_xml_file
        write_xml_declaration
        # Write the table element.
        write_table
        # Write the autoFilter element.
        write_auto_filter
        # Write the tableColumns element.
        write_table_columns
        # Write the tableStyleInfo element.
        write_table_style_info

        # Close the table tag
        @writer.end_tag('table')

        # Close the XML writer object and filehandle.
        @writer.crlf
        @writer.close
      end

      private

      #
      # Write the XML declaration.
      #
      def write_xml_declaration
        @writer.xml_decl('UTF-8', 1)
      end

      #
      # Write the <autoFilter> element.
      #
      def write_auto_filter
        autofilter = @properties[:_autofilter]

        return unless autofilter

        attributes = ['ref', autofilter]

        @writer.empty_tag('autoFilter', attributes)
      end

      #
      # Write the <table> element.
      #
      def write_table
        schema           = 'http://schemas.openxmlformats.org/'
        xmlns            = "#{schema}spreadsheetml/2006/main"
        id               = @properties[:id]
        name             = @properties[:_name]
        display_name     = @properties[:_name]
        ref              = @properties[:_range]
        totals_row_shown = @properties[:_totals_row_shown]
        header_row_count = @properties[:_header_row_count]

        attributes = [
                      'xmlns',       xmlns,
                      'id',          id,
                      'name',        name,
                      'displayName', display_name,
                      'ref',         ref
                     ]

        if header_row_count.nil? || header_row_count == 0
          attributes << 'headerRowCount' << 0
        end

        if totals_row_shown && totals_row_shown != 0
          attributes << 'totalsRowCount' << 1
        else
          attributes << 'totalsRowShown' << 0
        end
        @writer.start_tag('table', attributes)
      end

      #
      # Write the <autoFilter> element.
      #
      def write_auto_filter
        autofilter = @properties[:_autofilter]

        return if autofilter.nil? || autofilter == 0

        attributes = ['ref', autofilter]

        @writer.empty_tag('autoFilter', attributes)
      end

      #
      # Write the <tableColumns> element.
      #
      def write_table_columns
        columns = @properties[:_columns]

        count = columns.size

        attributes = ['count', count]

        @writer.tag_elements('tableColumns', attributes) do
          columns.each {|col_data| write_table_column(col_data)}
        end
      end

      #
      # Write the <tableColumn> element.
      #
      def write_table_column(col_data)
        attributes = [
                      'id',   col_data[:_id],
                      'name', col_data[:_name]
                     ]

        if col_data[:_total_string] && col_data[:_total_string] != ''
          attributes << :totalsRowLabel << col_data[:_total_string]
        elsif col_data[:_total_function] && col_data[:_total_function] != ''
          attributes << :totalsRowFunction << col_data[:_total_function]
        end

        if col_data[:_format]
          attributes << :dataDxfId << col_data[:_format]
        end

        if col_data[:_formula] && col_data[:_formula] != ''
          @writer.tag_elements('tableColumn', attributes) do
            # Write the calculatedColumnFormula element.
            write_calculated_column_formula(col_data[:_formula])
          end
        else
          @writer.empty_tag('tableColumn', attributes)
        end
      end

      #
      # Write the <tableStyleInfo> element.
      #
      def write_table_style_info
        props = @properties

        name                = props[:_style]
        show_first_column   = props[:_show_first_col]
        show_last_column    = props[:_show_last_col]
        show_row_stripes    = props[:_show_row_stripes]
        show_column_stripes = props[:_show_col_stripes]

        attributes = [
                      'name',              name,
                      'showFirstColumn',   show_first_column,
                      'showLastColumn',    show_last_column,
                      'showRowStripes',    show_row_stripes,
                      'showColumnStripes', show_column_stripes
                     ]

        @writer.empty_tag('tableStyleInfo', attributes)
      end

      #
      # Write the <calculatedColumnFormula> element.
      #
      def write_calculated_column_formula(formula)
        @writer.data_element('calculatedColumnFormula', formula)
      end
    end
  end
end
