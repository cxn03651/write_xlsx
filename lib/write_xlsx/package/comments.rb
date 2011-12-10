# -*- coding: utf-8 -*-
require 'write_xlsx/package/xml_writer_simple'
require 'write_xlsx/utility'

module Writexlsx
  module Package
    class Comments

      include Utility

      def initialize
        @writer = Package::XMLWriterSimple.new
        @author_ids = {}
      end

      def set_xml_writer(filename)
        @writer.set_xml_writer(filename)
      end

      def assemble_xml_file(comments_data)
        write_xml_declaration
        write_comments
        write_authors(comments_data)
        write_comment_list(comments_data)

        @writer.end_tag('comments')
        @writer.crlf
        @writer.close
      end

      private

      def write_xml_declaration
        @writer.xml_decl
      end

      #
      # Write the <comments> element.
      #
      def write_comments
        xmlns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

        attributes = [ 'xmlns', xmlns]

        @writer.start_tag('comments', attributes)
      end

      #
      # Write the <authors> element.
      #
      def write_authors(comment_data)
        author_count = 0

        @writer.start_tag('authors')
        comment_data.each do |comment|
          author = comment[3] || ''

          if author && !@author_ids[author]
            # Store the author id.
            @author_ids[author] = author_count
            author_count += 1

            # Write the author element.
            write_author(author)
          end
        end

        @writer.end_tag('authors')
      end

      #
      # Write the <author> element.
      #
      def write_author(data)
        @writer.data_element('author', data)
      end

      #
      # Write the <commentList> element.
      #
      def write_comment_list(comment_data)
        @writer.start_tag('commentList')

        comment_data.each do |comment|
          row    = comment[0]
          col    = comment[1]
          text   = comment[2]
          author = comment[3]

          # Look up the author id.
          author_id = nil
          author_id = @author_ids[author] if author

          # Write the comment element.
          write_comment(row, col, text, author_id)
        end

        @writer.end_tag( 'commentList' )
      end

      #
      # Write the <comment> element.
      #
      def write_comment(row, col, text, author_id)
        ref       = xl_rowcol_to_cell( row, col )
        author_id ||= 0

        attributes = ['ref', ref]

        (attributes << 'authorId' << author_id ) if author_id

        @writer.start_tag('comment', attributes)
        write_text(text)
        @writer.end_tag('comment')
      end

      #
      # Write the <text> element.
      #
      def write_text(text)
        @writer.start_tag('text')

        # Write the text r element.
        write_text_r(text)

        @writer.end_tag('text')
      end

      #
      # Write the <r> element.
      #
      def write_text_r(text)
        @writer.start_tag('r')

        # Write the rPr element.
        write_r_pr

        # Write the text r element.
        write_text_t(text)

        @writer.end_tag('r')
      end

      #
      # Write the text <t> element.
      #
      def write_text_t(text)
        attributes = []

        (attributes << 'xml:space' << 'preserve') if text =~ /^\s/ || text =~ /\s$/

        @writer.data_element('t', text, attributes)
      end

      #
      # Write the <rPr> element.
      #
      def write_r_pr
        @writer.start_tag('rPr')

        # Write the sz element.
        write_sz

        # Write the color element.
        write_color

        # Write the rFont element.
        write_r_font

        # Write the family element.
        write_family

        @writer.end_tag('rPr')
      end

      #
      # Write the <sz> element.
      #
      def write_sz
        val  = 8

        attributes = ['val', val]

        @writer.empty_tag('sz', attributes)
      end

      #
      # Write the <color> element.
      #
      def write_color
        indexed = 81

        attributes = ['indexed', indexed]

        @writer.empty_tag('color', attributes)
      end

      #
      # Write the <rFont> element.
      #
      def write_r_font
        val  = 'Tahoma'

        attributes = ['val', val]

        @writer.empty_tag('rFont', attributes)
      end

      #
      # Write the <family> element.
      #
      def write_family
        val  = 2

        attributes = ['val', val]

        @writer.empty_tag('family', attributes)
      end
    end
  end
end
