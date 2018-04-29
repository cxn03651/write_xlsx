# -*- coding: utf-8 -*-
require 'write_xlsx/col_name'

module Writexlsx
  module Utility
    ROW_MAX  = 1048576  # :nodoc:
    COL_MAX  = 16384    # :nodoc:
    STR_MAX  = 32767    # :nodoc:
    SHEETNAME_MAX = 31  # :nodoc:

    #
    # xl_rowcol_to_cell($row, col, row_absolute, col_absolute)
    #
    def xl_rowcol_to_cell(row, col, row_absolute = false, col_absolute = false)
      row += 1      # Change from 0-indexed to 1 indexed.
      col_str = xl_col_to_name(col, col_absolute)
      "#{col_str}#{absolute_char(row_absolute)}#{row}"
    end

    #
    # Returns: [row, col, row_absolute, col_absolute]
    #
    # The row_absolute and col_absolute parameters aren't documented because they
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
      col_str = ColName.instance.col_str(col)
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

    #
    # Sheetnames used in references should be quoted if they contain any spaces,
    # special characters or if the look like something that isn't a sheet name.
    # TODO. We need to handle more special cases.
    #
    def quote_sheetname(sheetname) #:nodoc:
      # Use Excel's conventions and quote the sheet name if it comtains any
      # non-word character or if it isn't already quoted.
      name = sheetname.dup
      if name =~ /\W/ && !(name =~ /^'/)
        # Double quote and single quoted strings.
        name = name.gsub(/'/, "''")
        name = "'#{name}'"
      end
      name
    end

    def check_dimensions(row, col)
      if !row || row >= ROW_MAX || !col || col >= COL_MAX
        raise WriteXLSXDimensionError
      end
      0
    end

    #
    # convert_date_time(date_time_string)
    #
    # The function takes a date and time in ISO8601 "yyyy-mm-ddThh:mm:ss.ss" format
    # and converts it to a decimal number representing a valid Excel date.
    #
    # Dates and times in Excel are represented by real numbers. The integer part of
    # the number stores the number of days since the epoch and the fractional part
    # stores the percentage of the day in seconds. The epoch can be either 1900 or
    # 1904.
    #
    # Parameter: Date and time string in one of the following formats:
    #               yyyy-mm-ddThh:mm:ss.ss  # Standard
    #               yyyy-mm-ddT             # Date only
    #                         Thh:mm:ss.ss  # Time only
    #
    # Returns:
    #            A decimal number representing a valid Excel date, or
    #            nil if the date is invalid.
    #
    def convert_date_time(date_time_string)       #:nodoc:
      date_time = date_time_string.sub(/^\s+/, '').sub(/\s+$/, '').sub(/Z$/, '')

      # Check for invalid date char.
      return nil if date_time =~ /[^0-9T:\-\.Z]/

      # Check for "T" after date or before time.
      return nil unless date_time =~ /\dT|T\d/

      days      = 0 # Number of days since epoch
      seconds   = 0 # Time expressed as fraction of 24h hours in seconds

      # Split into date and time.
      date, time = date_time.split(/T/)

      # We allow the time portion of the input DateTime to be optional.
      if time
        # Match hh:mm:ss.sss+ where the seconds are optional
        if time =~ /^(\d\d):(\d\d)(:(\d\d(\.\d+)?))?/
          hour   = $1.to_i
          min    = $2.to_i
          sec    = $4.to_f || 0
        else
          return nil # Not a valid time format.
        end

        # Some boundary checks
        return nil if hour >= 24
        return nil if min  >= 60
        return nil if sec  >= 60

        # Excel expresses seconds as a fraction of the number in 24 hours.
        seconds = (hour * 60* 60 + min * 60 + sec) / (24.0 * 60 * 60)
      end

      # We allow the date portion of the input DateTime to be optional.
      return seconds if date == ''

      # Match date as yyyy-mm-dd.
      if date =~ /^(\d\d\d\d)-(\d\d)-(\d\d)$/
        year   = $1.to_i
        month  = $2.to_i
        day    = $3.to_i
      else
        return nil  # Not a valid date format.
      end

      # Set the epoch as 1900 or 1904. Defaults to 1900.
      # Special cases for Excel.
      unless date_1904?
        return      seconds if date == '1899-12-31' # Excel 1900 epoch
        return      seconds if date == '1900-01-00' # Excel 1900 epoch
        return 60 + seconds if date == '1900-02-29' # Excel false leapday
      end


      # We calculate the date by calculating the number of days since the epoch
      # and adjust for the number of leap days. We calculate the number of leap
      # days by normalising the year in relation to the epoch. Thus the year 2000
      # becomes 100 for 4 and 100 year leapdays and 400 for 400 year leapdays.
      #
      epoch   = date_1904? ? 1904 : 1900
      offset  = date_1904? ?    4 :    0
      norm    = 300
      range   = year - epoch

      # Set month days and check for leap year.
      mdays   = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
      leap    = 0
      leap    = 1  if year % 4 == 0 && year % 100 != 0 || year % 400 == 0
      mdays[1]   = 29 if leap != 0

      # Some boundary checks
      return nil if year  < epoch or year  > 9999
      return nil if month < 1     or month > 12
      return nil if day   < 1     or day   > mdays[month - 1]

      # Accumulate the number of days since the epoch.
      days = day                               # Add days for current month
      (0 .. month-2).each do |m|
        days += mdays[m]                      # Add days for past months
      end
      days += range * 365                      # Add days for past years
      days += ((range)                /  4)    # Add leapdays
      days -= ((range + offset)       /100)    # Subtract 100 year leapdays
      days += ((range + offset + norm)/400)    # Add 400 year leapdays
      days -= leap                             # Already counted above

      # Adjust for Excel erroneously treating 1900 as a leap year.
      days += 1 if !date_1904? and days > 59

      date_time = sprintf("%0.10f", days + seconds)
      date_time = date_time.sub(/\.?0+$/, '') if date_time =~ /\./
      if date_time =~ /\./
        date_time.to_f
      else
        date_time.to_i
      end
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

    # Check for a cell reference in A1 notation and substitute row and column
    def row_col_notation(args)   # :nodoc:
      if args[0] =~ /^\D/
        substitute_cellref(*args)
      else
        args
      end
    end

    #
    # Substitute an Excel cell reference in A1 notation for  zero based row and
    # column values in an argument list.
    #
    # Ex: ("A4", "Hello") is converted to (3, 0, "Hello").
    #
    def substitute_cellref(cell, *args)       #:nodoc:
      return [*args] if cell.respond_to?(:coerce) # Numeric

      normalized_cell = cell.upcase

      case normalized_cell
      # Convert a column range: 'A:A' or 'B:G'.
      # A range such as A:A is equivalent to A1:65536, so add rows as required
      when /\$?([A-Z]{1,3}):\$?([A-Z]{1,3})/
        row1, col1 =  xl_cell_to_rowcol($1 + '1')
        row2, col2 =  xl_cell_to_rowcol($2 + ROW_MAX.to_s)
        return [row1, col1, row2, col2, *args]
      # Convert a cell range: 'A1:B7'
      when /\$?([A-Z]{1,3}\$?\d+):\$?([A-Z]{1,3}\$?\d+)/
        row1, col1 =  xl_cell_to_rowcol($1)
        row2, col2 =  xl_cell_to_rowcol($2)
        return [row1, col1, row2, col2, *args]
      # Convert a cell reference: 'A1' or 'AD2000'
      when /\$?([A-Z]{1,3}\$?\d+)/
        row1, col1 =  xl_cell_to_rowcol($1)
        return [row1, col1, *args]
      else
        raise("Unknown cell reference #{normalized_cell}")
      end
    end

    def underline_attributes(underline)
      if underline == 2
        [['val', 'double']]
      elsif underline == 33
        [['val', 'singleAccounting']]
      elsif underline == 34
        [['val', 'doubleAccounting']]
      else
        []    # Default to single underline.
      end
    end

    #
    # Write the <color> element.
    #
    def write_color(writer, name, value) #:nodoc:
      attributes = [[name, value]]

      writer.empty_tag('color', attributes)
    end

    #
    # return perl's boolean result
    #
    def ptrue?(value)
      if [false, nil, 0, "0", "", [], {}].include?(value)
        false
      else
        true
      end
    end

    def check_parameter(params, valid_keys, method)
      invalids = params.keys - valid_keys
      unless invalids.empty?
        raise WriteXLSXOptionParameterError,
          "Unknown parameter '#{invalids.join(', ')}' in #{method}."
      end
      true
    end

    #
    # Check that row and col are valid and store max and min values for use in
    # other methods/elements.
    #
    # The ignore_row/ignore_col flags is used to indicate that we wish to
    # perform the dimension check without storing the value.
    #
    # The ignore flags are use by set_row() and data_validate.
    #
    def check_dimensions_and_update_max_min_values(row, col, ignore_row = 0, ignore_col = 0)       #:nodoc:
      check_dimensions(row, col)
      store_row_max_min_values(row) if ignore_row == 0
      store_col_max_min_values(col) if ignore_col == 0

      0
    end

    def store_row_max_min_values(row)
      @dim_rowmin = row if !@dim_rowmin || (row < @dim_rowmin)
      @dim_rowmax = row if !@dim_rowmax || (row > @dim_rowmax)
    end

    def store_col_max_min_values(col)
      @dim_colmin = col if !@dim_colmin || (col < @dim_colmin)
      @dim_colmax = col if !@dim_colmax || (col > @dim_colmax)
    end

    def float_to_str(float)
      return '' unless float
      if float == float.to_i
        float.to_i.to_s
      else
        float.to_s
      end
    end

    #
    # Convert user defined layout properties to the format required internally.
    #
    def layout_properties(args, is_text = false)
      return unless ptrue?(args)

      properties = is_text ? [:x, :y] : [:x, :y, :width, :height]

      # Check for valid properties.
      args.keys.each do |key|
        unless properties.include?(key.to_sym)
            raise "Property '#{key}' not allowed in layout options\n"
        end
      end

      # Set the layout properties
      layout = Hash.new
      properties.each do |property|
        value = args[property]
        # Convert to the format used by Excel for easier testing.
        layout[property] = sprintf("%.17g", value)
      end

      layout
    end

    #
    # Convert vertices from pixels to points.
    #
    def pixels_to_points(vertices)
      col_start, row_start, x1,    y1,
      col_end,   row_end,   x2,    y2,
      left,      top,       width, height  = vertices.flatten

      left   *= 0.75
      top    *= 0.75
      width  *= 0.75
      height *= 0.75

      [left, top, width, height]
    end

    def v_shape_attributes_base(id, z_index)
      [
       ['id',          "_x0000_s#{id}"],
       ['type',        type],
       ['style',       (v_shape_style_base(z_index, vertices) + style_addition).join]
      ]
    end

    def v_shape_style_base(z_index, vertices)
      left, top, width, height = pixels_to_points(vertices)

      left_str    = float_to_str(left)
      top_str     = float_to_str(top)
      width_str   = float_to_str(width)
      height_str  = float_to_str(height)
      z_index_str = float_to_str(z_index)

      shape_style_base(left_str, top_str, width_str, height_str, z_index_str)
    end

    def shape_style_base(left_str, top_str, width_str, height_str, z_index_str)
      [
       'position:absolute;',
       'margin-left:',
       left_str, 'pt;',
       'margin-top:',
       top_str, 'pt;',
       'width:',
       width_str, 'pt;',
       'height:',
       height_str, 'pt;',
       'z-index:',
       z_index_str, ';'
      ]
    end

    #
    # Write the <v:fill> element.
    #
    def write_fill
      @writer.empty_tag('v:fill', fill_attributes)
    end

    #
    # Write the <v:path> element.
    #
    def write_comment_path(gradientshapeok, connecttype)
      attributes      = []

      attributes << ['gradientshapeok', 't'] if gradientshapeok
      attributes << ['o:connecttype', connecttype]

      @writer.empty_tag('v:path', attributes)
    end

    #
    # Write the <x:Anchor> element.
    #
    def write_anchor
      col_start, row_start, x1, y1, col_end, row_end, x2, y2 = vertices
      data = [col_start, x1, row_start, y1, col_end, x2, row_end, y2].join(', ')

      @writer.data_element('x:Anchor', data)
    end

    #
    # Write the <x:AutoFill> element.
    #
    def write_auto_fill
      @writer.data_element('x:AutoFill', 'False')
    end

    #
    # Write the <div> element.
    #
    def write_div(align, font = nil)
      style = "text-align:#{align}"
      attributes = [['style', style]]

      @writer.tag_elements('div', attributes) do
        if font
          # Write the font element.
          write_font(font)
        end
      end
    end

    #
    # Write the <font> element.
    #
    def write_font(font)
      caption = font[:_caption]
      face    = 'Calibri'
      size    = 220
      color   = '#000000'

      attributes = [
                    ['face',  face],
                    ['size',  size],
                    ['color', color]
                   ]
      @writer.data_element('font', caption, attributes)
    end

    #
    # Write the <v:stroke> element.
    #
    def write_stroke
      attributes = [['joinstyle', 'miter']]

      @writer.empty_tag('v:stroke', attributes)
    end

    def r_id_attributes(id)
      ['r:id', "rId#{id}"]
    end

    def write_xml_declaration
      @writer.xml_decl
      yield
      @writer.crlf
      @writer.close
    end

    #
    # Convert user defined line properties to the structure required internally.
    #
    def line_properties(line) # :nodoc:
      line_fill_properties(line) do
        value_or_raise(dash_types, line[:dash_type], 'dash type')
      end
    end

    #
    # Convert user defined fill properties to the structure required internally.
    #
    def fill_properties(fill) # :nodoc:
      line_fill_properties(fill)
    end

    def line_fill_properties(params)
      return { :_defined => 0 } unless params
      ret = params.dup
      ret[:dash_type] = yield if block_given? && ret[:dash_type]
      ret[:_defined] = 1
      ret
    end

    def dash_types
      {
        :solid               => 'solid',
        :round_dot           => 'sysDot',
        :square_dot          => 'sysDash',
        :dash                => 'dash',
        :dash_dot            => 'dashDot',
        :long_dash           => 'lgDash',
        :long_dash_dot       => 'lgDashDot',
        :long_dash_dot_dot   => 'lgDashDotDot',
        :dot                 => 'dot',
        :system_dash_dot     => 'sysDashDot',
        :system_dash_dot_dot => 'sysDashDotDot'
      }
    end

    def value_or_raise(hash, key, msg)
      raise "Unknown #{msg} '#{key}'" if hash[key.to_sym].nil?
      hash[key.to_sym]
    end

    def palette_color(index)
      # Adjust the colour index.
      idx = index - 8

      rgb = @palette[idx]
      sprintf("%02X%02X%02X", *rgb)
    end

    #
    # Workbook の生成時のオプションハッシュを解析する
    #
    def process_workbook_options(*params)
      case params.size
      when 0
        [{}, {}]
      when 1 # one hash
        options_keys = [:tempdir, :date_1904, :optimization, :excel2003_style, :strings_to_urls]

        hash = params.first
        options = hash.reject{|k,v| !options_keys.include?(k)}

        default_format_properties =
          hash[:default_format_properties] ||
          hash.reject{|k,v| options_keys.include?(k)}

        [options, default_format_properties.dup]
      when 2 # array which includes options and default_format_properties
        options, default_format_properties = params
        default_format_properties ||= {}

        [options.dup, default_format_properties.dup]
      end
    end
  end

  module WriteDPtPoint
      #
      # Write an individual <c:dPt> element. Override the parent method to add
      # markers.
      #
      def write_d_pt_point(index, point)
        @writer.tag_elements('c:dPt') do
          # Write the c:idx element.
          write_idx(index)
          @writer.tag_elements('c:marker') do
            # Write the c:spPr element.
            write_sp_pr(point)
          end
        end
      end
  end
end
