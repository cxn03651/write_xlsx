# -*- coding: utf-8 -*-
require 'write_xlsx/format'
require 'write_xlsx/package/xml_writer_simple'
require 'write_xlsx/utility'

module Writexlsx
  module Package

    class Comment

      include Writexlsx::Utility

      DEFAULT_COLOR  = 81  # what color ?
      DEFAULT_WIDTH  = 128
      DEFAULT_HEIGHT = 74

      attr_reader :row, :col, :string, :color, :vertices
      attr_accessor :author, :visible

      def initialize(workbook, worksheet, row, col, string, options = {})
        options ||= {}
        @workbook   = workbook
        @worksheet  = worksheet
        @row        = row
        @col        = col
        @string     = string[0, STR_MAX]
        @author     = options[:author]
        @color      = backgrount_color(options[:color] || DEFAULT_COLOR)
        @start_cell = options[:start_cell]
        @start_row, @start_col = if @start_cell
          substitute_cellref(@start_cell)
        else
          [ options[:start_row], options[:start_col] ]
        end
        @start_row  ||= default_start_row(row)
        @start_col  ||= default_start_col(col)
        @visible    = options[:visible]
        @x_offset   = options[:x_offset] || default_x_offset(col)
        @y_offset   = options[:y_offset] || default_y_offset(row)
        @x_scale    = options[:x_scale]  || 1
        @y_scale    = options[:y_scale]  || 1
        @width      = (0.5 + (options[:width]  || DEFAULT_WIDTH)  * @x_scale).to_i
        @height     = (0.5 + (options[:height] || DEFAULT_HEIGHT) * @y_scale).to_i
        @vertices   = @worksheet.position_object_pixels(
                                             @start_col, @start_row, @x_offset, @y_offset,
                                             @width, @height
                                             ) << [@width, @height]
      end

      def backgrount_color(color)
        color_id = Format.get_color(color)

        if color_id == 0
          @color = '#ffffe1'
        else
          rgb = @workbook.palette[color_id - 8]
          @color = "##{rgb_color(rgb)} [#{color_id}]\n"
        end
      end

      # Minor modification to allow comparison testing. Change RGB colors
      # from long format, ffcc00 to short format fc0 used by VML.
      def rgb_color(rgb)
        result = sprintf("%02x%02x%02x", *rgb)
        if result =~ /^([0-9a-f])\1([0-9a-f])\2([0-9a-f])\3$/
          result = "#{$1}#{$2}#{$3}"
        end
        result
      end

      def default_start_row(row)
        case row
        when 0
          0
        when ROW_MAX - 3
          ROW_MAX - 7
        when ROW_MAX - 2
          ROW_MAX - 6
        when ROW_MAX - 1
          ROW_MAX - 5
        else
          row - 1
        end
      end

      def default_start_col(col)
        case col
        when COL_MAX - 3
          COL_MAX - 6
        when COL_MAX - 2
          COL_MAX - 5
        when COL_MAX - 1
          COL_MAX - 4
        else
          col + 1
        end
      end

      def default_x_offset(col)
        case col
        when COL_MAX - 3, COL_MAX - 2, COL_MAX - 1
          49
        else
          15
        end
      end

      def default_y_offset(row)
        case row
        when 0
          2
        when ROW_MAX - 3, ROW_MAX - 2
          16
        when ROW_MAX - 1
          14
        else
          10
        end
      end
    end

    class Comments

      include Writexlsx::Utility

      def initialize(worksheet)
        @worksheet = worksheet
        @writer = Package::XMLWriterSimple.new
        @author_ids = {}
        @comments = {}
      end

      def add(comment)
        if @comments[comment.row]
          @comments[comment.row][comment.col] = comment
        else
          @comments[comment.row] = {}
          @comments[comment.row][comment.col] = comment
        end
      end

      def empty?
        @comments.empty?
      end

      def set_xml_writer(filename)
        @writer.set_xml_writer(filename)
      end

      def assemble_xml_file
        write_xml_declaration
        write_comments
        write_authors(sorted_comments)
        write_comment_list(sorted_comments)

        @writer.end_tag('comments')
        @writer.crlf
        @writer.close
      end

      def sorted_comments
        comments = []

        # We sort the comments by row and column but that isn't strictly required.
        @comments.keys.sort.each do |row|
          @comments[row].keys.sort.each do |col|
            # Set comment visibility if required and not already user defined.
            @comments[row][col].visible ||= 1 if comments_visible?

            # Set comment author if not already user defined.
            @comments[row][col].author ||= @comments_author
            comments << @comments[row][col]
          end
        end
        comments
      end

      def has_comment_in_row?(row)
        !!@comments[row]
      end

      private

      def comments_visible?
        @worksheet.comments_visible?
      end

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
          author = comment.author || ''
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
        comment_data.each { |comment| write_comment(comment) }
        @writer.end_tag( 'commentList' )
      end

      #
      # Write the <comment> element.
      #
      def write_comment(comment)
        ref       = xl_rowcol_to_cell( comment.row, comment.col )
        attributes = ['ref', ref]

        author_id = (@author_ids[comment.author] if comment.author) || 0
        attributes << 'authorId' << author_id

        @writer.start_tag('comment', attributes)
        write_text(comment.string)
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
