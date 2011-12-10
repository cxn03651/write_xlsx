# -*- coding: utf-8 -*-
require 'write_xlsx/package/xml_writer_simple'
require 'write_xlsx/utility'

module Writexlsx
  module Package
    class Core

      include Utility

      App_package  = 'application/vnd.openxmlformats-package.'
      App_document = 'application/vnd.openxmlformats-officedocument.'

      def initialize
        @writer = Package::XMLWriterSimple.new
        @properties = {}
        @localtime  = [Time.now]
      end

      def set_xml_writer(filename)
        @writer.set_xml_writer(filename)
      end

      def assemble_xml_file
        write_xml_declaration
        write_cp_core_properties
        write_dc_title
        write_dc_subject
        write_dc_creator
        write_cp_keywords
        write_dc_description
        write_cp_last_modified_by
        write_dcterms_created
        write_dcterms_modified
        write_cp_category
        write_cp_content_status

        @writer.end_tag('cp:coreProperties')
        @writer.crlf
        @writer.close
      end

      def set_properties(properties)
        @properties = properties
      end

      private

      #
      # Convert a localtime() date to a ISO 8601 style "2010-01-01T00:00:00Z" date.
      #
      def localtime_to_iso8601_date(local_time = nil)
        local_time ||= Time.now

        date = local_time.strftime('%Y-%m-%dT%H:%M:%SZ')
      end

      def write_xml_declaration
        @writer.xml_decl
      end

      #
      # Write the <cp:coreProperties> element.
      #
      def write_cp_core_properties
        xmlns_cp       = 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties'
        xmlns_dc       = 'http://purl.org/dc/elements/1.1/'
        xmlns_dcterms  = 'http://purl.org/dc/terms/'
        xmlns_dcmitype = 'http://purl.org/dc/dcmitype/'
        xmlns_xsi      = 'http://www.w3.org/2001/XMLSchema-instance'

        attributes = [
            'xmlns:cp',       xmlns_cp,
            'xmlns:dc',       xmlns_dc,
            'xmlns:dcterms',  xmlns_dcterms,
            'xmlns:dcmitype', xmlns_dcmitype,
            'xmlns:xsi',      xmlns_xsi
        ]

        @writer.start_tag('cp:coreProperties', attributes)
      end

      #
      # Write the <dc:creator> element.
      #
      def write_dc_creator
        data = @properties[:author] || ''

        @writer.data_element('dc:creator', data)
      end

      #
      # Write the <cp:lastModifiedBy> element.
      #
      def write_cp_last_modified_by
        data = @properties[:author] || ''

        @writer.data_element('cp:lastModifiedBy', data)
      end

      #
      # Write the <dcterms:created> element.
      #
      def write_dcterms_created
        date     = @properties[:created]
        xsi_type = 'dcterms:W3CDTF'

        date = localtime_to_iso8601_date(date)

        attributes = ['xsi:type', xsi_type]

        @writer.data_element('dcterms:created', date, attributes)
      end

      #
      # Write the <dcterms:modified> element.
      #
      def write_dcterms_modified
        date     = @properties[:created]
        xsi_type = 'dcterms:W3CDTF'

        date =  localtime_to_iso8601_date(date)

        attributes = ['xsi:type', xsi_type]

        @writer.data_element('dcterms:modified', date, attributes)
      end

      #
      # Write the <dc:title> element.
      #
      def write_dc_title
        data = @properties[:title]

        return unless data

        @writer.data_element('dc:title', data)
      end

      #
      # Write the <dc:subject> element.
      #
      def write_dc_subject
        data = @properties[:subject]

        return unless data

        @writer.data_element('dc:subject', data)
      end

      #
      # Write the <cp:keywords> element.
      #
      def write_cp_keywords
        data = @properties[:keywords]

        return unless data

        @writer.data_element('cp:keywords', data)
      end

      #
      # Write the <dc:description> element.
      #
      def write_dc_description
        data = @properties[:comments]

        return unless data

        @writer.data_element('dc:description', data)
      end

      #
      # Write the <cp:category> element.
      #
      def write_cp_category
        data = @properties[:category]

        return unless data

        @writer.data_element('cp:category', data)
      end

      #
      # Write the <cp:contentStatus> element.
      #
      def write_cp_content_status
        data = @properties[:status]

        return unless data

        @writer.data_element('cp:contentStatus', data)
      end
    end
  end
end
