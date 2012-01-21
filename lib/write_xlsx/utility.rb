# -*- coding: utf-8 -*-
module Writexlsx
  module Utility
    ROW_MAX  = 1048576  # :nodoc:
    COL_MAX  = 16384    # :nodoc:
    STR_MAX  = 32767    # :nodoc:

    #
    # xl_rowcol_to_cell($row, $col, $row_absolute, $col_absolute)
    #
    def xl_rowcol_to_cell(row, col, row_absolute = false, col_absolute = false)
      row += 1      # Change from 0-indexed to 1 indexed.
      row_abs = row_absolute ? '$' : ''
      col_abs = col_absolute ? '$' : ''
      col_str = xl_col_to_name(col, col_absolute)
      "#{col_str}#{absolute_char(row_absolute)}#{row}"
    end

    #
    # Returns: ($row, $col, $row_absolute, $col_absolute)
    #
    # The $row_absolute and $col_absolute parameters aren't documented because they
    # mainly used internally and aren't very useful to the user.
    #
    def xl_cell_to_rowcol(cell)
      cell =~ /(\$?)([A-Z]{1,3})(\$?)(\d+)/

      col_abs = $1 != ""
      col     = $2
      row_abs = $3 != ""
      row     = $4.to_i

      # Convert base26 column string to number
      # All your Base are belong to us.
      chars = col.split(//)
      expn = 0
      col = 0

      chars.reverse.each do |char|
        col += (char.ord - 'A'.ord + 1) * (26 ** expn)
        expn += 1
      end

      # Convert 1-index to zero-index
      row -= 1
      col -= 1

      return [row, col, row_abs, col_abs]
    end

    def xl_col_to_name(col, col_absolute)
      # Change from 0-indexed to 1 indexed.
      col += 1
      col_str = ''

      while col > 0
        # Set remainder from 1 .. 26
        remainder = col % 26
        remainder = 26 if remainder == 0

        # Convert the remainder to a character. C-ishly.
        col_letter = ("A".ord + remainder - 1).chr

        # Accumulate the column letters, right to left.
        col_str = col_letter + col_str

        # Get the next order of magnitude.
        col = (col - 1) / 26
      end

      "#{absolute_char(col_absolute)}#{col_str}"
    end

    def xl_range(row_1, row_2, col_1, col_2,
                 row_abs_1 = false, row_abs_2 = false, col_abs_1 = false, col_abs_2 = false)
      range1 = xl_rowcol_to_cell(row_1, col_1, row_abs_1, col_abs_1)
      range2 = xl_rowcol_to_cell(row_2, col_2, row_abs_2, col_abs_2)

      "#{range1}:#{range2}"
    end

    def xl_range_formula(sheetname, row_1, row_2, col_1, col_2)
      # Use Excel's conventions and quote the sheet name if it contains any
      # non-word character or if it isn't already quoted.
      sheetname = "'#{sheetname}'" if sheetname =~ /\W/ && !(sheetname =~ /^'/)

      range1 = xl_rowcol_to_cell( row_1, col_1, 1, 1 )
      range2 = xl_rowcol_to_cell( row_2, col_2, 1, 1 )

      "=#{sheetname}!#{range1}:#{range2}"
    end

    def absolute_char(absolute)
      absolute ? '$' : ''
    end

    def xml_str
      @writer.string
    end
    
    def self.delete_files(path)
      if FileTest.file?(path)
        File.delete(path)
      elsif FileTest.directory?(path)
        Dir.foreach(path) do |file|
          next if file =~ /^\.\.?$/  # '.' or '..'
          delete_files(path.sub(/\/+$/,"") + '/' + file)
        end
        Dir.rmdir(path)
      end
    end

    def put_deprecate_message(method)
      $stderr.puts("Warning: calling deprecated method #{method}. This method will be removed in a future release.")
    end
  end
end
