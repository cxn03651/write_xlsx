# -*- coding: utf-8 -*-
require 'write_xlsx/package/xml_writer_simple'
require 'write_xlsx/utility'

module Writexlsx
  module Package
    class Vml

      include Writexlsx::Utility

      def initialize
        @writer = Package::XMLWriterSimple.new
      end

      def set_xml_writer(filename)
        @writer.set_xml_writer(filename)
      end

      def assemble_xml_file(worksheet)
        return unless @writer

        write_xml_namespace do
          # Write the o:shapelayout element.
          write_shapelayout(worksheet.vml_data_id)

          z_index = 1
          vml_shape_id = worksheet.vml_shape_id
          unless worksheet.buttons_data.empty?
            vml_shape_id, z_index =
              write_shape_type_and_shape(
                                         worksheet.buttons_data,
                                         vml_shape_id, z_index) do
              write_button_shapetype
            end
          end
          unless worksheet.sorted_comments.empty?
            write_shape_type_and_shape(
                                       worksheet.sorted_comments,
                                       vml_shape_id, z_index) do
              write_comment_shapetype
            end
          end
        end
        @writer.crlf
        @writer.close
      end

      private

      def write_shape_type_and_shape(data, vml_shape_id, z_index)
        # Write the v:shapetype element.
        yield
        data.each do |obj|
          # Write the v:shape element.
          vml_shape_id += 1
          obj.write_shape(@writer, vml_shape_id, z_index)
          z_index += 1
        end
        [vml_shape_id, z_index]
      end

      #
      # Write the <xml> element. This is the root element of VML.
      #
      def write_xml_namespace
        @writer.tag_elements('xml', xml_attributes) do
          yield
        end
      end

      # for <xml> elements.
      def xml_attributes
        schema  = 'urn:schemas-microsoft-com:'
        [
         'xmlns:v', "#{schema}vml",
         'xmlns:o', "#{schema}office:office",
         'xmlns:x', "#{schema}office:excel"
        ]
      end

      #
      # Write the <o:shapelayout> element.
      #
      def write_shapelayout(data_id)
        attributes = ['v:ext', 'edit']

        @writer.tag_elements('o:shapelayout', attributes) do
          # Write the o:idmap element.
          write_idmap(data_id)
        end
      end

      #
      # Write the <o:idmap> element.
      #
      def write_idmap(data_id)
        attributes = [
          'v:ext', 'edit',
          'data',  data_id
        ]

        @writer.empty_tag('o:idmap', attributes)
      end

      #
      # Write the <v:shapetype> element.
      #
      def write_comment_shapetype
        attributes = [
            'id',        '_x0000_t202',
            'coordsize', '21600,21600',
            'o:spt',     202,
            'path',      'm,l,21600r21600,l21600,xe'
        ]

        @writer.tag_elements('v:shapetype', attributes) do
          # Write the v:stroke element.
          write_stroke
          # Write the v:path element.
          write_comment_path('t', 'rect')
        end
      end

      #
      # Write the <v:shapetype> element.
      #
      def write_button_shapetype
        attributes = [
                      'id',        '_x0000_t201',
                      'coordsize', '21600,21600',
                      'o:spt',     201,
                      'path',      'm,l,21600r21600,l21600,xe'
                     ]

        @writer.tag_elements('v:shapetype', attributes) do
          # Write the v:stroke element.
          write_stroke
          # Write the v:path element.
          write_button_path
          # Write the o:lock element.
          write_shapetype_lock
        end
      end

      #
      # Write the <v:path> element.
      #
      def write_button_path
        attributes = [
                      'shadowok',      'f',
                      'o:extrusionok', 'f',
                      'strokeok',      'f',
                      'fillok',        'f',
                      'o:connecttype', 'rect'
                     ]
        @writer.empty_tag('v:path', attributes)
      end

      #
      # Write the <o:lock> element.
      #
      def write_shapetype_lock
        attributes = [
                      'v:ext',     'edit',
                      'shapetype', 't'
                     ]
        @writer.empty_tag('o:lock', attributes)
      end
    end
  end
end
