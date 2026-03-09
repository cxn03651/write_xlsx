# -*- coding: utf-8 -*-
# frozen_string_literal: true

###############################################################################
#
# Formatting - A module for worksheet layout and print/appearance settings.
#
# Used in conjunction with WriteXLSX
#
# Copyright 2000-2011, John McNamara, jmcnamara@cpan.org
# Convert to ruby by Hideo NAKAMURA, nakamura.hideo@gmail.com

module Writexlsx
  class Worksheet
    module Formatting
      include Utility

      #
      # Header/footer
      #

      #
      # Set the page header caption and optional margin.
      #
      def set_header(string = '', margin = 0.3, options = {})
        raise 'Header string must be less than 255 characters' if string.length > 255

        # Replace the Excel placeholder &[Picture] with the internal &G.
        header_footer_string = string.gsub("&[Picture]", '&G')
        # placeholeder /&G/ の数
        placeholder_count = header_footer_string.scan("&G").count
        @page_setup.header = header_footer_string

        @page_setup.header_footer_aligns = options[:align_with_margins] if options[:align_with_margins]

        @page_setup.header_footer_scales = options[:scale_with_doc] if options[:scale_with_doc]

        # Reset the array in case the function is called more than once.
        @header_images = []

        [
          [:image_left, 'LH'], [:image_center, 'CH'], [:image_right, 'RH']
        ].each do |p|
          @header_images << ImageProperty.new(options[p.first], position: p.last) if options[p.first]
        end

        # # placeholeder /&G/ の数
        # placeholder_count = @page_setup.header.scan("&G").count

        raise "Number of header image (#{@header_images.size}) doesn't match placeholder count (#{placeholder_count}) in string: #{@page_setup.header}" if @header_images.size != placeholder_count

        @page_setup.margin_header         = margin || 0.3
        @page_setup.header_footer_changed = true
      end

      #
      # Set the page footer caption and optional margin.
      #
      def set_footer(string = '', margin = 0.3, options = {})
        raise 'Footer string must be less than 255 characters' if string.length > 255

        # Replace the Excel placeholder &[Picture] with the internal &G.
        @page_setup.footer = string.gsub("&[Picture]", '&G')

        @page_setup.header_footer_aligns = options[:align_with_margins] if options[:align_with_margins]

        @page_setup.header_footer_scales = options[:scale_with_doc] if options[:scale_with_doc]

        # Reset the array in case the function is called more than once.
        @footer_images = []

        [
          [:image_left, 'LF'], [:image_center, 'CF'], [:image_right, 'RF']
        ].each do |p|
          @footer_images << ImageProperty.new(options[p.first], position: p.last) if options[p.first]
        end

        # placeholeder /&G/ の数
        placeholder_count = @page_setup.footer.scan("&G").count

        raise "Number of footer image (#{@footer_images.size}) doesn't match placeholder count (#{placeholder_count}) in string: #{@page_setup.footer}" if @footer_images.size != placeholder_count

        @page_setup.margin_footer         = margin
        @page_setup.header_footer_changed = true
      end

      #
      # Page margin helpers
      #

      #
      # Center the worksheet data horizontally between the margins on the printed page:
      #
      def center_horizontally
        @page_setup.center_horizontally
      end

      #
      # Center the worksheet data vertically between the margins on the printed page:
      #
      def center_vertically
        @page_setup.center_vertically
      end

      #
      # Set all the page margins to the same value in inches.
      #
      def margins=(margin)
        self.margin_left   = margin
        self.margin_right  = margin
        self.margin_top    = margin
        self.margin_bottom = margin
      end

      #
      # Set the left and right margins to the same value in inches.
      # See set_margins
      #
      def margins_left_right=(margin)
        self.margin_left  = margin
        self.margin_right = margin
      end

      #
      # Set the top and bottom margins to the same value in inches.
      # See set_margins
      #
      def margins_top_bottom=(margin)
        self.margin_top    = margin
        self.margin_bottom = margin
      end

      #
      # Set the left margin in inches.
      # See margins=()
      #
      def margin_left=(margin)
        @page_setup.margin_left = remove_white_space(margin)
      end

      #
      # Set the right margin in inches.
      # See margins=()
      #
      def margin_right=(margin)
        @page_setup.margin_right = remove_white_space(margin)
      end

      #
      # Set the top margin in inches.
      # See margins=()
      #
      def margin_top=(margin)
        @page_setup.margin_top = remove_white_space(margin)
      end

      #
      # Set the bottom margin in inches.
      # See margins=()
      #
      def margin_bottom=(margin)
        @page_setup.margin_bottom = remove_white_space(margin)
      end

      # # deprecations for set_* wrapper methods
      #
      # set_margin_* methods are deprecated. use margin_*=().
      #
      def set_margins(margin)
        put_deprecate_message("#{self}.set_margins")
        self.margins = margin
      end

      #
      # this method is deprecated. use margin_left_right=().
      # Set the left and right margins to the same value in inches.
      #
      def set_margins_LR(margin)
        put_deprecate_message("#{self}.set_margins_LR")
        self.margins_left_right = margin
      end

      #
      # this method is deprecated. use margin_top_bottom=().
      # Set the top and bottom margins to the same value in inches.
      #
      def set_margins_TB(margin)
        put_deprecate_message("#{self}.set_margins_TB")
        self.margins_top_bottom = margin
      end

      #
      # this method is deprecated. use margin_left=()
      # Set the left margin in inches.
      #
      def set_margin_left(margin = 0.7)
        put_deprecate_message("#{self}.set_margin_left")
        self.margin_left = margin
      end

      #
      # this method is deprecated. use margin_right=()
      # Set the right margin in inches.
      #
      def set_margin_right(margin = 0.7)
        put_deprecate_message("#{self}.set_margin_right")
        self.margin_right = margin
      end

      #
      # this method is deprecated. use margin_top=()
      # Set the top margin in inches.
      #
      def set_margin_top(margin = 0.75)
        put_deprecate_message("#{self}.set_margin_top")
        self.margin_top = margin
      end

      #
      # this method is deprecated. use margin_bottom=()
      # Set the bottom margin in inches.
      #
      def set_margin_bottom(margin = 0.75)
        put_deprecate_message("#{self}.set_margin_bottom")
        self.margin_bottom = margin
      end

      #
      # Repeat/print area
      #

      #
      # Set the number of rows to repeat at the top of each printed page.
      #
      def repeat_rows(row_min, row_max = nil)
        row_max ||= row_min

        # Convert to 1 based.
        row_min += 1
        row_max += 1

        area = "$#{row_min}:$#{row_max}"

        # Build up the print titles "Sheet1!$1:$2"
        sheetname = quote_sheetname(@name)
        @page_setup.repeat_rows = "#{sheetname}!#{area}"
      end

      def print_repeat_rows   # :nodoc:
        @page_setup.repeat_rows
      end

      #
      # :call-seq:
      #   repeat_columns(first_col, last_col = nil)
      #
      # Set the columns to repeat at the left hand side of each printed page.
      #
      def repeat_columns(*args)
        if args[0] =~ /^\D/
          _dummy, first_col, _dummy, last_col = substitute_cellref(*args)
        else
          first_col, last_col = args
        end
        last_col ||= first_col

        area = "#{xl_col_to_name(first_col, 1)}:#{xl_col_to_name(last_col, 1)}"
        @page_setup.repeat_cols = "#{quote_sheetname(@name)}!#{area}"
      end

      def print_repeat_cols  # :nodoc:
        @page_setup.repeat_cols
      end

      #
      # :call-seq:
      #   print_area(first_row, first_col, last_row, last_col)
      #
      # This method is used to specify the area of the worksheet that will
      # be printed. All four parameters must be specified. You can also use
      # A1 notation.
      #
      def print_area(*args)
        return @page_setup.print_area.dup if args.empty?

        if (row_col_array = row_col_notation(args.first))
          row1, col1, row2, col2 = row_col_array
        else
          row1, col1, row2, col2 = args
        end

        return if [row1, col1, row2, col2].include?(nil)

        # Ignore max print area since this is the same as no print area for Excel.
        return if row1 == 0 && col1 == 0 && row2 == ROW_MAX - 1 && col2 == COL_MAX - 1

        # Build up the print area range "=Sheet2!R1C1:R2C1"
        @page_setup.print_area = convert_name_area(row1, col1, row2, col2)
      end

      #
      # Scale and view
      #

      #
      # Set the worksheet zoom factor in the range <tt>10 <= scale <= 400</tt>:
      #
      def zoom=(scale)
        # Confine the scale to Excel's range
        @zoom = if scale < 10 || scale > 400
                  # carp "Zoom factor scale outside range: 10 <= zoom <= 400"
                  100
                else
                  scale.to_i
                end
      end

      # This method is deprecated. use zoom=().
      def set_zoom(scale)
        put_deprecate_message("#{self}.set_zoom")
        self.zoom = scale
      end

      #
      # Set the scale factor of the printed page.
      # Scale factors in the range 10 <= scale <= 400 are valid:
      #
      def print_scale=(scale = 100)
        scale_val = scale.to_i
        # Confine the scale to Excel's range
        scale_val = 100 if scale_val < 10 || scale_val > 400

        # Turn off "fit to page" option.
        @page_setup.fit_page = false

        @page_setup.scale              = scale_val
        @page_setup.page_setup_changed = true
      end

      #
      # This method is deprecated. use print_scale=().
      #
      def set_print_scale(scale = 100)
        put_deprecate_message("#{self}.set_print_scale")
        self.print_scale = (scale)
      end

      #
      # Set the option to print the worksheet in black and white.
      #
      def print_black_and_white
        @page_setup.black_white        = true
        @page_setup.page_setup_changed = true
      end

      #
      # Display the worksheet right to left for some eastern versions of Excel.
      #
      def right_to_left(flag = true)
        @right_to_left = !!flag
      end

      #
      # Hide cell zero values.
      #
      def hide_zero(flag = true)
        @show_zeros = !flag
      end

      #
      # Set the paper type. Ex. 1 = US Letter, 9 = A4
      #
      def paper=(paper_size)
        @page_setup.paper = paper_size
      end

      def set_paper(paper_size)
        put_deprecate_message("#{self}.set_paper")
        self.paper = paper_size
      end

      #
      # Set the order in which pages are printed.
      #
      def print_across(across = true)
        if across
          @page_setup.across             = true
          @page_setup.page_setup_changed = true
        else
          @page_setup.across = false
        end
      end

      #
      # The start_page=() method is used to set the number of the
      # starting page when the worksheet is printed out.
      #
      def start_page=(page_start)
        @page_setup.page_start = page_start
      end

      def set_start_page(page_start)
        put_deprecate_message("#{self}.set_start_page")
        self.start_page = page_start
      end
    end
  end
end
