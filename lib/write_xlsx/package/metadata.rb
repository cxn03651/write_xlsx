# -*- coding: utf-8 -*-
# frozen_string_literal: true

require 'write_xlsx/package/xml_writer_simple'
require 'write_xlsx/utility'

module Writexlsx
  module Package
    #
    # Metadata - A class for writing the Excel XLSX metadata.xml file.
    #
    class Metadata
      include Writexlsx::Utility

      def initialize(workbook)
        @writer = Package::XMLWriterSimple.new
        @workbook      = workbook
      end

      def set_xml_writer(filename)
        @writer.set_xml_writer(filename)
      end

      def assemble_xml_file
        write_xml_declaration do
          # Write the metadata element.
          write_metadata
          # Write the metadataTypes element.
          write_metadata_types
          # Write the futureMetadata element.
          write_future_metadata
          # Write the cellMetadata element.
          write_cell_metadata
          @writer.end_tag('metadata')
        end
      end

      private

      #
      # Write the <metadata> element.
      #
      def write_metadata
        xmlns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
        xmlns_xda =
          'http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray'

        attributes = [
          ['xmlns',     xmlns],
          ['xmlns:xda', xmlns_xda]
        ]

        @writer.start_tag('metadata', attributes)
      end

      #
      # Write the <metadataTypes> element.
      #
      def write_metadata_types
        attributes = [['count', 1]]

        @writer.tag_elements('metadataTypes', attributes) do
          # Write the metadataType element.
          write_metadata_type
        end
      end

      #
      # Write the <metadataType> element.
      #
      def write_metadata_type
        attributes = [
          %w[name XLDAPR],
          ['minSupportedVersion', 120000],
          ['copy',                1],
          ['pasteAll',            1],
          ['pasteValues',         1],
          ['merge',               1],
          ['splitFirst',          1],
          ['rowColShift',         1],
          ['clearFormats',        1],
          ['clearComments',       1],
          ['assign',              1],
          ['coerce',              1],
          ['cellMeta',            1]
        ]

        @writer.empty_tag('metadataType', attributes)
      end

      #
      # Write the <futureMetadata> element.
      #
      def write_future_metadata
        attributes = [
          %w[name XLDAPR],
          ['count', 1]
        ]

        @writer.tag_elements('futureMetadata', attributes) do
          @writer.tag_elements('bk') do
            @writer.tag_elements('extLst') do
              # Write the ext element.
              write_ext
            end
          end
        end
      end

      #
      # Write the <ext> element.
      #
      def write_ext
        attributes = [['uri', '{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}']]
        @writer.tag_elements('ext', attributes) do
          # Write the xda:dynamicArrayProperties element.
          write_xda_dynamic_array_properties
        end
      end

      #
      # Write the <xda:dynamicArrayProperties> element.
      #
      def write_xda_dynamic_array_properties
        attributes = [
          ['fDynamic',   1],
          ['fCollapsed', 0]
        ]

        @writer.empty_tag('xda:dynamicArrayProperties', attributes)
      end

      #
      # Write the <cellMetadata> element.
      #
      def write_cell_metadata
        count = 1

        attributes = [['count', count]]

        @writer.tag_elements('cellMetadata', attributes) do
          @writer.tag_elements('bk') do
            # Write the rc element.
            write_rc
          end
        end
      end

      #
      # Write the <rc> element.
      #
      def write_rc
        attributes = [
          ['t', 1],
          ['v', 0]
        ]
        @writer.empty_tag('rc', attributes)
      end
    end
  end
end
