# -*- coding: utf-8 -*-
# frozen_string_literal: true

module Writexlsx
  module Utility
    module StringWidth
      MAX_DIGIT_WIDTH    = 7    # For Calabri 11.  # :nodoc:
      PADDING            = 5                       # :nodoc:
      DEFAULT_COL_PIXELS = 64
      CHAR_WIDTHS = {
        ' '  =>  3, '!' =>  5, '"' =>  6, '#' =>  7, '$' =>  7, '%' => 11,
        '&'  => 10, "'" =>  3, '(' =>  5, ')' =>  5, '*' =>  7, '+' =>  7,
        ','  =>  4, '-' =>  5, '.' =>  4, '/' =>  6, '0' =>  7, '1' =>  7,
        '2'  =>  7, '3' =>  7, '4' =>  7, '5' =>  7, '6' =>  7, '7' =>  7,
        '8'  =>  7, '9' =>  7, ':' =>  4, ';' =>  4, '<' =>  7, '=' =>  7,
        '>'  =>  7, '?' =>  7, '@' => 13, 'A' =>  9, 'B' =>  8, 'C' =>  8,
        'D'  =>  9, 'E' =>  7, 'F' =>  7, 'G' =>  9, 'H' =>  9, 'I' =>  4,
        'J'  =>  5, 'K' =>  8, 'L' =>  6, 'M' => 12, 'N' => 10, 'O' => 10,
        'P'  =>  8, 'Q' => 10, 'R' =>  8, 'S' =>  7, 'T' =>  7, 'U' =>  9,
        'V'  =>  9, 'W' => 13, 'X' =>  8, 'Y' =>  7, 'Z' =>  7, '[' =>  5,
        '\\' =>  6, ']' =>  5, '^' =>  7, '_' =>  7, '`' =>  4, 'a' =>  7,
        'b'  =>  8, 'c' =>  6, 'd' =>  8, 'e' =>  8, 'f' =>  5, 'g' =>  7,
        'h'  =>  8, 'i' =>  4, 'j' =>  4, 'k' =>  7, 'l' =>  4, 'm' => 12,
        'n'  =>  8, 'o' =>  8, 'p' =>  8, 'q' =>  8, 'r' =>  5, 's' =>  6,
        't'  =>  5, 'u' =>  8, 'v' =>  7, 'w' => 11, 'x' =>  7, 'y' =>  7,
        'z'  =>  6, '{' =>  5, '|' =>  7, '}' =>  5, '~' =>  7
      }.freeze

      #
      # xl_string_pixel_width($string)
      #
      # Get the pixel width of a string based on individual character widths taken
      # from Excel. UTF8 characters are given a default width of 8.
      #
      # Note, Excel adds an additional 7 pixels padding to a cell.
      #
      def xl_string_pixel_width(string)
        length = 0
        string.to_s.chars.each { |char| length += CHAR_WIDTHS[char] || 8 }

        length
      end
    end
  end
end
