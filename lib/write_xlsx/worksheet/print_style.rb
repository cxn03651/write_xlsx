# -*- encoding: utf-8 -*-

module Writexlsx
  class Worksheet
    class PrintStyle # :nodoc:
      attr_accessor :margin_left, :margin_right, :margin_top, :margin_bottom  # :nodoc:
      attr_accessor :margin_header, :margin_footer                            # :nodoc:
      attr_accessor :repeat_rows, :repeat_cols, :print_area                   # :nodoc:
      attr_accessor :hbreaks, :vbreaks, :scale                                # :nodoc:
      attr_accessor :fit_page, :fit_width, :fit_height, :page_setup_changed   # :nodoc:
      attr_accessor :across                                                   # :nodoc:
      attr_accessor :orientation  # :nodoc:

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

      def attributes    # :nodoc:
        [
         'left',   @margin_left,
         'right',  @margin_right,
         'top',    @margin_top,
         'bottom', @margin_bottom,
         'header', @margin_header,
         'footer', @margin_footer
        ]
      end

      def orientation?
        !!@orientation
      end
    end
  end
end
