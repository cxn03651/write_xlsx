# -*- encoding: utf-8 -*-

module Writexlsx
  class Worksheet
    class PageSetup # :nodoc:
      attr_accessor :margin_left, :margin_right, :margin_top, :margin_bottom  # :nodoc:
      attr_accessor :margin_header, :margin_footer                            # :nodoc:
      attr_accessor :repeat_rows, :repeat_cols, :print_area                   # :nodoc:
      attr_accessor :hbreaks, :vbreaks, :scale                                # :nodoc:
      attr_accessor :fit_page, :fit_width, :fit_height, :page_setup_changed   # :nodoc:
      attr_writer :across                                                   # :nodoc:
      attr_accessor :orientation, :print_options_changed  # :nodoc:
      attr_accessor :header, :footer, :header_footer_changed

      def initialize # :nodoc:
        @margin_left = 0.7
        @margin_right = 0.7
        @margin_top = 0.75
        @margin_bottom = 0.75
        @margin_header = 0.3
        @margin_footer = 0.3
        @repeat_rows   = ''
        @repeat_cols   = ''
        @print_area    = ''
        @hbreaks = []
        @vbreaks = []
        @scale = 100
        @fit_page = false
        @fit_width  = nil
        @fit_height = nil
        @page_setup_changed = false
        @across = false
        @orientation = true
      end

      def paper=(paper_size)
        if paper_size
          @paper_size = paper_size
          @page_setup_changed = true
        end
      end

      def center_horizontally
        @print_options_changed = true
        @hcenter               = true
      end

      def center_vertically
        @print_options_changed = true
        @vcenter               = true
      end

      def print_row_col_headers(headers)
        if headers
          @print_headers         = true
          @print_options_changed = true
        else
          @print_headers         = false
        end
      end

      def hide_gridlines(option)
        if option == 0 || !option
          @print_gridlines  = true
          @print_options_changed = true
        elsif option = 1
          @print_gridlines  = false
        else
          @print_gridlines  = false
        end
      end

      #
      # Write the <pageSetup> element.
      #
      # The following is an example taken from Excel.
      #
      # <pageSetup
      #     paperSize="9"
      #     scale="110"
      #     fitToWidth="2"
      #     fitToHeight="2"
      #     pageOrder="overThenDown"
      #     orientation="portrait"
      #     blackAndWhite="1"
      #     draft="1"
      #     horizontalDpi="200"
      #     verticalDpi="200"
      #     r:id="rId1"
      # />
      #
      def write_page_setup(writer) #:nodoc:
        return unless @page_setup_changed

        attributes = []
        attributes << 'paperSize'   << @paper_size    if @paper_size
        attributes << 'scale'       << @scale         if @scale != 100
        attributes << 'fitToWidth'  << @fit_width     if @fit_page && @fit_width != 1
        attributes << 'fitToHeight' << @fit_height    if @fit_page && @fit_height != 1
        attributes << 'pageOrder'   << "overThenDown" if @across
        attributes << 'orientation'
        if @orientation
          attributes << 'portrait'
        else
          attributes << 'landscape'
        end

        writer.empty_tag('pageSetup', attributes)
      end

      #
      # Write the <pageMargins> element.
      #
      def write_page_margins(writer) #:nodoc:
        writer.empty_tag('pageMargins', margin_attributes)
      end

      #
      # Write the <printOptions> element.
      #
      def write_print_options(writer) #:nodoc:
        return unless @print_options_changed

        attributes = []
        attributes << 'horizontalCentered' << 1 if @hcenter
        attributes << 'verticalCentered'   << 1 if @vcenter
        attributes << 'headings'           << 1 if @print_headers
        attributes << 'gridLines'          << 1 if @print_gridlines
        writer.empty_tag('printOptions', attributes)
      end

      #
      # Write the <headerFooter> element.
      #
      def write_header_footer(writer) #:nodoc:
        return unless @header_footer_changed

        writer.tag_elements('headerFooter') do
          write_odd_header(writer) if @header && @header != ''
          write_odd_footer(writer) if @footer && @footer != ''
        end
      end

      private

      #
      # Write the <oddHeader> element.
      #
      def write_odd_header(writer) #:nodoc:
        writer.data_element('oddHeader', @header)
      end

      #
      # Write the <oddFooter> element.
      #
      def write_odd_footer(writer) #:nodoc:
        writer.data_element('oddFooter', @footer)
      end

      def margin_attributes    # :nodoc:
        [
         'left',   @margin_left,
         'right',  @margin_right,
         'top',    @margin_top,
         'bottom', @margin_bottom,
         'header', @margin_header,
         'footer', @margin_footer
        ]
      end
    end
  end
end
