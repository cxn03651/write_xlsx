# -*- coding: utf-8 -*-
require 'write_xlsx/package/xml_writer_simple'
require 'write_xlsx/colors'
require 'write_xlsx/format'
require 'write_xlsx/drawing'
require 'write_xlsx/compatibility'
require 'write_xlsx/utility'
require 'tempfile'

module Writexlsx
  class Worksheet
    include Writexlsx::Utility

    RowMax   = 1048576  # :nodoc:
    ColMax   = 16384    # :nodoc:
    StrMax   = 32767    # :nodoc:
    Buffer   = 4096     # :nodoc:

    attr_writer :fit_page
    attr_reader :print_area, :index, :_repeat_cols, :_repeat_rows
    attr_reader :charts, :images, :drawing
    attr_reader :external_hyper_links, :external_drawing_links, :external_comment_links, :drawing_links
    attr_reader :vml_data_id, :vml_shape_id, :comments_array
    attr_reader :autofilter_area, :hidden

    def initialize(workbook, index, name)
      @writer = Package::XMLWriterSimple.new

      @workbook = workbook
      @index = index
      @name = name
      @colinfo = []
      @table = []
      @filter_on = false
      @margin_left = 0.7
      @margin_right = 0.7
      @margin_top = 0.75
      @margin_bottom = 0.75
      @margin_header = 0.3
      @margin_footer = 0.3
      @_repeat_rows   = ''
      @_repeat_cols   = ''
      @print_area    = ''
      @screen_gridlines = true
      @show_zeros = true
      @xls_rowmax = RowMax
      @xls_colmax = ColMax
      @xls_strmax = StrMax
      @dim_rowmin = nil
      @dim_rowmax = nil
      @dim_colmin = nil
      @dim_colmax = nil
      @selections = []
      @panes = []

      @tab_color  = 0

      @orientation = true

      @hbreaks = []
      @vbreaks = []

      @set_cols = {}
      @set_rows = {}
      @zoom = 100
      @zoom_scale_normal = true
      @right_to_left = false

      @autofilter_area = nil
      @filter_on    = false
      @filter_range = []
      @filter_cols  = {}
      @filter_type  = {}

      @col_sizes = {}
      @row_sizes = {}
      @col_formats = {}

      @hlink_count            = 0
      @hlink_refs             = []
      @external_hyper_links   = []
      @external_drawing_links = []
      @external_comment_links = []
      @drawing_links          = []
      @charts                 = []
      @images                 = []

      @zoom = 100
      @print_scale = 100
      @outline_row_level = 0
      @outline_col_level = 0

      @merge = []

      @has_comments = false
      @comments = {}

      @fit_page = false

      @validations = []

      @cond_formats = {}
      @dxf_priority = 1
    end

    def set_xml_writer(filename)
      @writer.set_xml_writer(filename)
    end

    def assemble_xml_file
      @writer.xml_decl
      write_worksheet
      write_sheet_pr
      write_dimension
      write_sheet_views
      write_sheet_format_pr
      write_cols
      write_sheet_data
      write_sheet_protection
      write_auto_filter
      write_merge_cells
      write_conditional_formats
      write_data_validations
      write_hyperlinks
      write_print_options
      write_page_margins
      write_page_setup
      write_header_footer
      write_row_breaks
      write_col_breaks
      write_drawings
      write_legacy_drawing
      # write_ext_lst
      @writer.end_tag('worksheet')
      @writer.crlf
      @writer.close
    end

    def name
      @name
    end

    #
    # Set this worksheet as a selected worksheet, i.e. the worksheet has its tab
    # highlighted.
    #
    def select
      @hidden   = false  # Selected worksheet can't be hidden.
      @selected = true
    end

    #
    # Set this worksheet as the active worksheet, i.e. the worksheet that is
    # displayed when the workbook is opened. Also set it as selected.
    #
    def activate
      @hidden = false
      @selected = true
      @workbook.activesheet = @index
    end

    #
    # Hide this worksheet.
    #
    def hide
      @hidden = true
      @selected = false
      @workbook.activesheet = 0
      @workbook.firstsheet  = 0
    end

    def hidden?
      @hidden
    end

    #
    # Set this worksheet as the first visible sheet. This is necessary
    # when there are a large number of worksheets and the activated
    # worksheet is not visible on the screen.
    #
    def set_first_sheet
      @hidden = false
      @workbook.firstsheet = self
    end

    #
    # Set the worksheet protection flags to prevent modification of worksheet
    # objects.
    #
    def protect(password = nil, options = {})
      # Default values for objects that can be protected.
      defaults = {
        :sheet                 => true,
        :content               => false,
        :objects               => false,
        :scenarios             => false,
        :format_cells          => false,
        :format_columns        => false,
        :format_rows           => false,
        :insert_columns        => false,
        :insert_rows           => false,
        :insert_hyperlinks     => false,
        :delete_columns        => false,
        :delete_rows           => false,
        :select_locked_cells   => true,
        :sort                  => false,
        :autofilter            => false,
        :pivot_tables          => false,
        :select_unlocked_cells => true,
      }

      # Overwrite the defaults with user specified values.
      options.each do |k, v|
        if defaults.has_key?(k)
          defaults[k] = options[k]
        else
          raise "Unknown protection object: #{k}\n"
        end
      end

      # Set the password after the user defined values.
      defaults[:password] =
        sprintf("%X", encode_password(password)) if password && password != ''

      @protect = defaults
    end

    #
    # set_column($firstcol, $lastcol, $width, $format, $hidden, $level)
    #
    # Set the width of a single column or a range of columns.
    # See also: _write_col_info
    #
    def set_column(*args)
      # Check for a cell reference in A1 notation and substitute row and column
      if args[0] =~ /^\D/
        row1, firstcol, row2, lastcol, *data = substitute_cellref(*args)
        args.shift        # remove row1
        args[1..1] = []   # remove row2
      else
        firstcol, lastcol, *data = args
      end

      # Ensure at least firstcol, lastcol and width
      return unless firstcol && lastcol && !data.empty?

      # Assume second column is the same as first if 0. Avoids KB918419 bug.
      lastcol = firstcol if lastcol == 0

      # Ensure 2nd col is larger than first. Also for KB918419 bug.
      firstcol, lastcol = lastcol, firstcol if firstcol > lastcol

      width, format, hidden, level = data

      # Check that cols are valid and store max and min values with default row.
      # NOTE: The check shouldn't modify the row dimensions and should only modify
      #       the column dimensions in certain cases.
      ignore_row = 1
      ignore_col = 1
      ignore_col = 0 if format.respond_to?(:get_xf_index) # Column has a format.
      ignore_col = 0 if width && hidden && hidden != 0    # Column has a width but is hidden

      return -2 if check_dimensions(0, firstcol, ignore_row, ignore_col) != 0
      return -2 if check_dimensions(0, lastcol,  ignore_row, ignore_col) != 0

      # Set the limits for the outline levels (0 <= x <= 7).
      level = 0 unless level
      level = 0 if level < 0
      level = 7 if level > 7

      @outline_col_level = level if level > @outline_col_level

      # Store the column data.
      @colinfo.push([firstcol, lastcol] + data)

      # Store the column change to allow optimisations.
      @col_size_changed = 1

      # Store the col sizes for use when calculating image vertices taking
      # hidden columns into account. Also store the column formats.
      width  ||= 0                        # Ensure width isn't undef.
      width = 0 if hidden && hidden != 0  # Set width to zero if col is hidden

      (firstcol .. lastcol).each do |col|
        @col_sizes[col]   = width
        @col_formats[col] = format if format
      end
    end

    #
    # Set which cell or cells are selected in a worksheet.
    #
    def set_selection(*args)
      return if args.empty?

      args = row_col_notation(args)

      # There should be either 2 or 4 arguments.
      case args.size
      when 2
        # Single cell selection.
        active_cell = xl_rowcol_to_cell( args[0], args[1] )
        sqref = active_cell
      when 4
        # Range selection.
        active_cell = xl_rowcol_to_cell( args[0], args[1] )

        row_first, col_first, row_last, col_last = args

        # Swap last row/col for first row/col as necessary
        row_first, row_last = row_last, row_first if row_first > row_last
        col_first, col_last = col_last, col_first if col_first > col_last

        # If the first and last cell are the same write a single cell.
        if (row_first == row_last) && (col_first == col_last)
          sqref = active_cell
        else
          sqref = xl_range(row_first, col_first, row_last, col_last)
        end
      else
        # User supplied wrong number or arguments.
        return
      end

      # Selection isn't set for cell A1.
      return if sqref == 'A1'

      @selections = [ [ nil, active_cell, sqref ] ]
    end

    #
    # Set panes and mark them as frozen.
    #
    def freeze_panes(*args)
      return if args.empty?

      # Check for a cell reference in A1 notation and substitute row and column.
      args = row_col_notation(args)

      row      = args[0]
      col      = args[1] || 0
      top_row  = args[2] || row
      left_col = args[3] || col
      type     = args[4] || 0

      @panes   = [row, col, top_row, left_col, type ]
    end

    #
    # Set panes and mark them as split.
    #
    # Implementers note. The API for this method doesn't map well from the XLS
    # file format and isn't sufficient to describe all cases of split panes.
    # It should probably be something like:
    #
    #     split_panes( $y, $x, $top_row, $left_col, $offset_row, $offset_col )
    #
    # I'll look at changing this if it becomes an issue.
    #
    def split_panes(*args)
      # Call freeze panes but add the type flag for split panes.
      freeze_panes(args[0], args[1], args[2], args[3], 2)
    end

    #
    # Set the page orientation as portrait.
    #
    def set_portrait
      @orientation        = true
      @page_setup_changed = true
    end

    #
    # Set the page orientation as landscape.
    #
    def set_landscape
      @orientation         = false
      @page_setup_changed  = true
    end

    #
    # Set the page view mode for Mac Excel.
    #
    def set_page_view(flag = true)
      @page_view = !!flag
    end

    #
    # Set the colour of the worksheet tab.
    #
    def set_tab_color(color)
      @tab_color = Colors.new.get_color(color)
    end

    #
    # Set the paper type. Ex. 1 = US Letter, 9 = A4
    #
    def set_paper(paper_size)
      if paper_size
        @paper_size         = paper_size
        @page_setup_changed = true
      end
    end

    #
    # Set the page header caption and optional margin.
    #
    def set_header(string = '', margin = 0.3)
      raise 'Header string must be less than 255 characters' if string.length >= 255

      @header                = string
      @margin_header         = margin
      @header_footer_changed = true
    end

    #
    # Set the page footer caption and optional margin.
    #
    def set_footer(string = '', margin = 0.3)
      raise 'Footer string must be less than 255 characters' if string.length >= 255

      @footer                = string
      @margin_footer         = margin
      @header_footer_changed = true
    end

    #
    # Center the page horizontally.
    #
    def center_horizontally
      @print_options_changed = true
      @hcenter               = true
    end

    #
    # Center the page horizontally.
    #
    def center_vertically
      @print_options_changed = true
      @vcenter               = true
    end

    #
    # Set all the page margins to the same value in inches.
    #
    def set_margins(margin)
      set_margin_left(margin)
      set_margin_right(margin)
      set_margin_top(margin)
      set_margin_bottom(margin)
    end

    #
    # Set the left and right margins to the same value in inches.
    #
    def set_margins_LR(margin)
      set_margin_left(margin)
      set_margin_right(margin)
    end

    #
    # Set the top and bottom margins to the same value in inches.
    #
    def set_margins_TB(margin)
      set_margin_top(margin)
      set_margin_bottom(margin)
    end

    #
    # Set the left margin in inches.
    #
    def set_margin_left(margin = 0.7)
      @margin_left = remove_white_space(margin)
    end

    #
    # Set the right margin in inches.
    #
    def set_margin_right(margin = 0.7)
      @margin_right = remove_white_space(margin)
    end

    #
    # Set the top margin in inches.
    #
    def set_margin_top(margin = 0.75)
      @margin_top = remove_white_space(margin)
    end

    #
    # Set the bottom margin in inches.
    #
    def set_margin_bottom(margin = 0.75)
      @margin_bottom = remove_white_space(margin)
    end

    #
    # Set the rows to repeat at the top of each printed page.
    #
    def repeat_rows(row_min, row_max = nil)
      row_max ||= row_min

      # Convert to 1 based.
      row_min += 1
      row_max += 1

      area = "$#{row_min}:$#{row_max}"

      # Build up the print titles "Sheet1!$1:$2"
      sheetname = quote_sheetname(name)
      @_repeat_rows = "#{sheetname}!#{area}"
    end

    #
    # Set the worksheet zoom factor.
    #
    def set_zoom(scale = 100)
      # Confine the scale to Excel's range
      if scale < 10 or scale > 400
        # carp "Zoom factor $scale outside range: 10 <= zoom <= 400"
        scale = 100
      end

      @zoom = scale.to_i
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
    # Set the order in which pages are printed.
    #
    def print_across(page_order = true)
      if page_order
        @page_order         = true
        @page_setup_changed = true
      else
        @page_order = false
      end
    end

    def write(*args)
      # Check for a cell reference in A1 notation and substitute row and column
      args = row_col_notation(args)

      token = args[2] || ''

      # Match an array ref.
      if token.respond_to?(:to_ary)
        write_row(*args)
      elsif token.respond_to?(:coerce)  # Numeric
        write_number(*args)
      elsif token =~ /^\d+$/
        write_number(*args)
      # Match http, https or ftp URL
      elsif token =~ %r|^[fh]tt?ps?://|
        write_url(*args)
      # Match mailto:
      elsif token =~ %r|^mailto:|
        write_url(*args)
      # Match internal or external sheet link
      elsif token =~ %r!^(?:in|ex)ternal:!
        write_url(*args)
      # Match formula
      elsif token =~ /^=/
        write_formula(*args)
      # Match array formula
      elsif token =~ /^\{=.*\}$/
        write_formula(*args)
      # Match blank
      elsif token == ''
        args.delete_at(2)     # remove the empty string from the parameter list
        write_blank(*args)
      else
        write_string(*args)
      end
    end

    #
    # write_row($row, $col, $array_ref, $format)
    #
    # Write a row of data starting from ($row, $col). Call write_col() if any of
    # the elements of the array ref are in turn array refs. This allows the writing
    # of 1D or 2D arrays of data in one go.
    #
    # Returns: the first encountered error value or zero for no errors
    #
    def write_row(*args)
      # Check for a cell reference in A1 notation and substitute row and column
      args = row_col_notation(args)

      # Catch non array refs passed by user.
      raise "Not an array ref in call to write_row()$!" unless args[2].respond_to?(:to_ary)

      row, col, tokens, *options = args
      error   = 0
      ret = 0

      tokens.each do |token|
        # Check for nested arrays
        if token.respond_to?(:to_ary)
          write_col(row, col, token, *options)
        else
          ret = write(row, col, token, *options)
        end

        # Return only the first error encountered, if any.
        error = ret if error == 0 && ret > 0
        col += 1
      end

      error
    end

    #
    # write_col($row, $col, $array_ref, $format)
    #
    # Write a column of data starting from ($row, $col). Call write_row() if any of
    # the elements of the array ref are in turn array refs. This allows the writing
    # of 1D or 2D arrays of data in one go.
    #
    # Returns: the first encountered error value or zero for no errors
    #
    def write_col(*args)
      # Check for a cell reference in A1 notation and substitute row and column
      args = row_col_notation(args)

      # Catch non array refs passed by user.
      raise "Not an array ref in call to write_col()$!" unless args[2].respond_to?(:to_ary)

      row, col, tokens, *options = args
      error   = 0
      ret = 0

      tokens.each do |token|
        # write() will deal with any nested arrays
        ret = write(row, col, token, *options)

        # Return only the first error encountered, if any.
        error = ret if error == 0 && ret > 0
        row += 1
      end

      error
    end

    #
    # Write a comment to the specified row and column (zero indexed).
    #
    # Returns  0 : normal termination
    #         -1 : insufficient number of arguments
    #         -2 : row or column out of range
    #
    def write_comment(*args)
      # Check for a cell reference in A1 notation and substitute row and column
      args = row_col_notation(args)

      return -1 if args.size < 3                # Check the number of args

      row, col = args

      # Check that row and col are valid and store max and min values
      return -2 if check_dimensions(row, col) != 0

      @has_comments = true
      # Process the properties of the cell comment.
      if @comments[row]
        @comments[row][col] = comment_params(*args)
      else
        @comments[row] = {}
        @comments[row][col] = comment_params(*args)
      end
    end

    def write_number(*args)
      # Check for a cell reference in A1 notation and substitute row and column
      args = row_col_notation(args)

      return -1 if args.size < 3                # Check the number of args
      row     = args[0]                         # Zero indexed row
      col     = args[1]                         # Zero indexed column
      num     = args[2]
      xf      = args[3]                         # The cell format
      type    = 'n'

      # Check that row and col are valid and store max and min values
      return -2 if check_dimensions(row, col) != 0

      store_data_to_table(row, col, [type, num, xf])
      0
    end

    #
    # Write a string to the specified row and column (zero indexed).
    # $format is optional.
    # Returns  0 : normal termination
    #         -1 : insufficient number of arguments
    #         -2 : row or column out of range
    #         -3 : long string truncated to 32767 chars
    #
    def write_string(*args)
      # Check for a cell reference in A1 notation and substitute row and column
      args = row_col_notation(args)

      return -1 if args.size < 3    # Check the number of args
      row, col, str, xf = args
      type = 's'                    # The data type

      # Check that row and col are valid and store max and min values
      return -2 if check_dimensions(row, col) != 0

      # Check that the string is < 32767 chars
      str_error = 0
      if str.length > @xls_strmax
        str = str[0, @xls_strmax]
        str_error = -3
      end

      index = shared_string_index(str)

      store_data_to_table(row, col, [type, index, xf])
      str_error
    end

    # write_rich_string( $row, $column, $format, $string, ..., $cell_format )
    #
    # The write_rich_string() method is used to write strings with multiple formats.
    # The method receives string fragments prefixed by format objects. The final
    # format object is used as the cell format.
    #
    # Returns  0 : normal termination.
    #         -1 : insufficient number of arguments.
    #         -2 : row or column out of range.
    #         -3 : long string truncated to 32767 chars.
    #         -4 : 2 consecutive formats used.
    #
    def write_rich_string(*args)
      # Check for a cell reference in A1 notation and substitute row and column
      args = row_col_notation(args)

      return -1 if args.size < 3     # Check the number of args

      row    = args.shift            # Zero indexed row.
      col    = args.shift            # Zero indexed column.
      str    = ''
      xf     = nil
      type   = 's'                   # The data type.
      length = 0                     # String length.

      # Check that row and col are valid and store max and min values
      return -2 if check_dimensions(row, col) != 0

      # If the last arg is a format we use it as the cell format.
      xf = args.pop if args[-1].respond_to?(:get_xf_index)

      # Create a temp XML::Writer object and use it to write the rich string
      # XML to a string.
      writer = Package::XMLWriterSimple.new

      @rstring = writer

      # Create a temp format with the default font for unformatted fragments.
      default = Format.new(0)

      # Convert the list of $format, $string tokens to pairs of ($format, $string)
      # except for the first $string fragment which doesn't require a default
      # formatting run. Use the default for strings without a leading format.
      last = 'format'
      pos  = 0

      fragments = []
      args.each do |token|
        if token.respond_to?(:get_xf_index)
          # Can't allow 2 formats in a row.
          return -4 if last == 'format' && pos > 0

          # Token is a format object. Add it to the fragment list.
          fragments << token
          last = 'format'
        else
          # Token is a string.
          if last != 'format'
            # If previous token wasn't a format add one before the string.
            fragments << default << token
          else
            # If previous token was a format just add the string.
            fragments << token
          end

          length += token.size    # Keep track of actual string length.
          last = 'string'
        end
        pos += 1
      end

      # If the first token is a string start the <r> element.
      @rstring.start_tag('r') if !fragments[0].respond_to?(:get_xf_index)

      # Write the XML elements for the $format $string fragments.
      fragments.each do |token|
        if token.respond_to?(:get_xf_index)
          # Write the font run.
          @rstring.start_tag('r')
          write_font(token)
        else
          # Write the string fragment part, with whitespace handling.
          attributes = []

          attributes << 'xml:space' << 'preserve' if token =~ /^\s/ || token =~ /\s$/
          @rstring.data_element('t', token, attributes)
          @rstring.end_tag('r')
        end
      end

      # Check that the string is < 32767 chars.
      return -3 if length > @xls_strmax

      # Add the XML string to the shared string table.
      index = get_shared_string_index(writer.string)

      store_data_to_table(row, col, [type, index, xf])

      return 0
    end

    #
    # write_blank($row, $col, $format)
    #
    # Write a blank cell to the specified row and column (zero indexed).
    # A blank cell is used to specify formatting without adding a string
    # or a number.
    #
    # A blank cell without a format serves no purpose. Therefore, we don't write
    # a BLANK record unless a format is specified. This is mainly an optimisation
    # for the write_row() and write_col() methods.
    #
    # Returns  0 : normal termination (including no format)
    #         -1 : insufficient number of arguments
    #         -2 : row or column out of range
    #
    def write_blank(*args)
      # Check for a cell reference in A1 notation and substitute row and column
      args = row_col_notation(args)

      # Check the number of args
      return -1 if args.size < 2

      # Don't write a blank cell unless it has a format
      return 0 unless args[2]

      row, col, xf = args
      type = 'b'                    # The data type

      # Check that row and col are valid and store max and min values
      return -2 if check_dimensions(row, col) != 0

      store_data_to_table(row, col, [type, nil, xf])
      0
    end

    def write_formula(*args)
      # Check for a cell reference in A1 notation and substitute row and column
      args = row_col_notation(args)

      return -1 if args.size < 3   # Check the number of args

      row, col, formula, format, value = args

      if formula =~ /^\{=.*\}$/
        return write_array_formula(row, col, row, col, formula, format, value )
      end

      # Check that row and col are valid and store max and min values
      return -2 unless check_dimensions(row, col) == 0

      formula.sub!(/^=/, '')

      store_data_to_table(row, col, ['f', formula, format, value])
      0
    end

    #
    # write_array_formula($row1, $col1, $row2, $col2, $formula, $format)
    #
    # Write an array formula to the specified row and column (zero indexed).
    #
    #  my $row1    = $_[0]           # First row
    #  my $col1    = $_[1]           # First column
    #  my $row2    = $_[2]           # Last row
    #  my $col2    = $_[3]           # Last column
    #  my $formula = $_[4]           # The formula text string
    #  my $xf      = $_[5]           # The format object.
    #  my $value   = $_[6]           # Optional formula value.
    #
    # $format is optional.
    #
    # Returns  0 : normal termination
    #         -1 : insufficient number of arguments
    #         -2 : row or column out of range
    #
    def write_array_formula(*args)
      # Check for a cell reference in A1 notation and substitute row and column
      args = row_col_notation(args)

      return -1 if args.size < 5    # Check the number of args

      row1, col1, row2, col2, formula, xf, value = args
      type = 'a'                    # The data type

      # Swap last row/col with first row/col as necessary
      row1, row2 = row2, row1 if row1 > row2
      col1, col2 = col1, col2 if col1 > col2

      # Check that row and col are valid and store max and min values
      return -2 if check_dimensions(row2, col2) != 0

      # Define array range
      if row1 == row2 && col1 == col2
        range = xl_rowcol_to_cell(row1, col1)
      else
        range ="#{xl_rowcol_to_cell(row1, col1)}:#{xl_rowcol_to_cell(row2, col2)}"
      end

      # Remove array formula braces and the leading =.
      formula.sub!(/^\{(.*)\}$/, '\1')
      formula.sub!(/^=/, '')

      store_data_to_table(row1, col1, [type, formula, xf, range, value])
      0
    end

    def store_formula(string)
      string.split(/(\$?[A-I]?[A-Z]\$?\d+)/)
    end

    #
    # write_url($row, $col, $url, $string, $format)
    #
    # Write a hyperlink. This is comprised of two elements: the visible label and
    # the invisible link. The visible label is the same as the link unless an
    # alternative string is specified. The label is written using the
    # write_string() method. Therefore the max characters string limit applies.
    # $string and $format are optional and their order is interchangeable.
    #
    # The hyperlink can be to a http, ftp, mail, internal sheet, or external
    # directory url.
    #
    # Returns  0 : normal termination
    #         -1 : insufficient number of arguments
    #         -2 : row or column out of range
    #         -3 : long string truncated to 32767 chars
    #
    def write_url(*args)
      # Check for a cell reference in A1 notation and substitute row and column
      args = row_col_notation(args)

      return -1 if args.size < 3    # Check the number of args

      # Reverse the order of $string and $format if necessary. We work on a copy
      # in order to protect the callers args. We don't use "local @_" in case of
      # perl50005 threads.
      args[3], args[4] = args[4], args[3] if args[3].respond_to?(:get_xf_index)

      row, col, url, str, xf, tip = args
      type      = 'l'                       # XML data type
      link_type = 1

      # Remove the URI scheme from internal links.
      if url =~ /^internal:/
        url.sub!(/^internal:/, '')
        link_type = 2
      # Remove the URI scheme from external links.
      elsif url =~ /^external:/
        url.sub!(/^external:/, '')
        link_type = 3
      end

      # The displayed string defaults to the url string.
      str = url unless str

      # For external links change the directory separator from Unix to Dos.
      if link_type == 3
        url.gsub!(%r|/|, '\\')
        str.gsub!(%r|/|, '\\')
      end

      # Strip the mailto header.
      str.sub!(/^mailto:/, '')

      # Check that row and col are valid and store max and min values
      return -2 if check_dimensions(row, col) != 0

      # Check that the string is < 32767 chars
      str_error = 0
      if str.bytesize > @xls_strmax
        str = str[0, @xls_strmax]
        str_error = -3
      end

      # Store the URL displayed text in the shared string table.
      index = get_shared_string_index(str)

      # External links to URLs and to other Excel workbooks have slightly
      # different characteristics that we have to account for.
      if link_type == 1
        # Ordinary URL style external links don't have a "location" string.
        str = nil
      elsif link_type == 3
        # External Workbook links need to be modified into the right format.
        # The URL will look something like 'c:\temp\file.xlsx#Sheet!A1'.
        # We need the part to the left of the # as the URL and the part to
        # the right as the "location" string (if it exists)
        url, str = url.split(/#/)

        # Add the file:/// URI to the url if non-local.
#            url = "file:///#{url}" if url =~ m{[\\/]} && url !~ m{^\.\.}


        # Treat as a default external link now that the data has been modified.
        link_type = 1
      end

      store_data_to_table(row, col, [type, index, xf, link_type, url, str, tip])

      return str_error
    end

    #
    # write_date_time ($row, $col, $string, $format)
    #
    # Write a datetime string in ISO8601 "yyyy-mm-ddThh:mm:ss.ss" format as a
    # number representing an Excel date. $format is optional.
    #
    # Returns  0 : normal termination
    #         -1 : insufficient number of arguments
    #         -2 : row or column out of range
    #         -3 : Invalid date_time, written as string
    #
    def write_date_time(*args)
      # Check for a cell reference in A1 notation and substitute row and column
      args = row_col_notation(args)

      return -1 if args.size < 3    # Check the number of args

      row, col, str, xf = args
      type = 'n'                    # The data type

      # Check that row and col are valid and store max and min values
      return -2 if check_dimensions(row, col) != 0

      str_error = 0
      date_time = convert_date_time(str)

      # If the date isn't valid then write it as a string.
      return write_string(args) unless date_time
      store_data_to_table(row, col, [type, date_time, xf])

      str_error
    end

    #
    # Insert a chart into a worksheet. The $chart argument should be a Chart
    # object or else it is assumed to be a filename of an external binary file.
    # The latter is for backwards compatibility.
    #
    def insert_chart(*args)
      # Check for a cell reference in A1 notation and substitute row and column.
      args = row_col_notation(args)
      return -1 if args.size < 3

      row, col, chart, x_offset, y_offset, scale_x, scale_y = args
      x_offset ||= 0
      y_offset ||= 0
      scale_x  ||= 1
      scale_y  ||= 1

      raise "Not a Chart object in insert_chart()" unless chart.is_a?(Chart)
      raise "Not a embedded style Chart object in insert_chart()" if chart.embedded == 0

      @charts << [row, col, chart, x_offset, y_offset, scale_x, scale_y]
    end

    #
    # Insert an image into the worksheet.
    #
    def insert_image(*args)
      # Check for a cell reference in A1 notation and substitute row and column.
      args = row_col_notation(args)
      return -1 if args.size < 3

      row, col, image, x_offset, y_offset, scale_x, scale_y = args
      x_offset ||= 0
      y_offset ||= 0
      scale_x  ||= 1
      scale_y  ||= 1

      @images << [row, col, image, x_offset, y_offset, scale_x, scale_y]
    end

    def repeat_formula(*args)
      args = row_col_notation(args)
      return -1 if args.size < 2   # Check the number of args
      row, col, formula, format, *pairs = args
      raise "Odd number of elements in pattern/replacement list" unless pairs.size % 2 == 0
      raise "Not a valid formula" unless formula.respond_to?(:to_ary)
      tokens  = formula.join("\t").split("\t")
      raise "No tokens in formula" if tokens.empty?
      value = nil
      if pairs[-2] == 'result'
        value = pairs.pop
        pairs.pop
      end
      while !pairs.empty?
        pattern = pairs.shift
        replace = pairs.shift

        tokens.each do |token|
          break if token.sub!(pattern, replace)
        end
      end
      formula = tokens.join('')
      write_formula(row, col, formula, format, value)
    end

    #
    # convert_date_time($date_time_string)
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
    #            undef if the date is invalid.
    #
    def convert_date_time(date_time_string)       #:nodoc:
      date_time = date_time_string

      days      = 0 # Number of days since epoch
      seconds   = 0 # Time expressed as fraction of 24h hours in seconds

      # Strip leading and trailing whitespace.
      date_time.sub!(/^\s+/, '')
      date_time.sub!(/\s+$/, '')

      # Check for invalid date char.
      return nil if date_time =~ /[^0-9T:\-\.Z]/

      # Check for "T" after date or before time.
      return nil unless date_time =~ /\dT|T\d/

      # Strip trailing Z in ISO8601 date.
      date_time.sub!(/Z$/, '')

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

      days + seconds
    end

    #
    # This method is used to set the height and XF format for a row.
    #
    def set_row(*args)
      row = args[0]
      height = args[1] || 15
      xf     = args[2]
      hidden = args[3] || 0
      level  = args[4] || 0
      collapsed = args[5] || 0

      return if row.nil?

      # Check that row and col are valid and store max and min values.
      return -2 if check_dimensions(row, 0) != 0

      # If the height is 0 the row is hidden and the height is the default.
      if height == 0
        hidden = 1
        height = 15
      end

      # Set the limits for the outline levels (0 <= x <= 7).
      level = 0 if level < 0
      level = 7 if level > 7

      @outline_row_level = level if level > @outline_row_level

      # Store the row properties.
      @set_rows[row] = [height, xf, hidden, level, collapsed]

      # Store the row change to allow optimisations.
      @row_size_changed = true

      # Store the row sizes for use when calculating image vertices.
      @row_sizes[row] = height
    end

    #
    # merge_range($first_row, $first_col, $last_row, $last_col, $string, $format)
    #
    # Merge a range of cells. The first cell should contain the data and the others
    # should be blank. All cells should contain the same format.
    #
    def merge_range(*args)
      # Check for a cell reference in A1 notation and substitute row and column
      args = row_col_notation(args)

      raise "Incorrect number of arguments" if args.size < 6
      raise "Fifth parameter must be a format object" unless args[5].kind_of?(Format)

      rwFirst, colFirst, rwLast, colLast, string, format, *extra_args = args

      # Excel doesn't allow a single cell to be merged
      raise "Can't merge single cell" if rwFirst == rwLast and colFirst == colLast

      # Swap last row/col with first row/col as necessary
      rwFirst,  rwLast  = rwLast,  rwFirst  if rwFirst > rwLast
      colFirst, colLast = colLast, colFirst if colFirst > colLast

      # Check that column number is valid and store the max value
      return if check_dimensions(rwLast, colLast) != 0

      # Store the merge range.
      @merge << [rwFirst, colFirst, rwLast, colLast]

      # Write the first cell
      write(rwFirst, colFirst, string, format, *extra_args)

      # Pad out the rest of the area with formatted blank cells.
      (rwFirst .. rwLast).each do |row|
        (colFirst .. colLast).each do |col|
          next if (row == rwFirst && col == colFirst)
          write_blank(row, col, format)
        end
      end
    end

    #
    # Same as merge_range() above except the type of write() is specified.
    #
    def merge_range_type(type, *args)
      # Check for a cell reference in A1 notation and substitute row and column
      args = row_col_notation(args)

      row_first = args.shift
      col_first = args.shift
      row_last  = args.shift
      col_last  = args.shift

      # Get the format. It can be in different positions for the different types.
      if type == 'array_formula' || type == 'blank' || type == 'rich_string'
        # The format is the last element.
        format = args[-1]
      else
        # Or else it is after the token.
        format = args[1]
      end

      # Check that there is a format object.
      raise "Format object missing or in an incorrect position" unless format.respond_to?(:get_xf_index)

      # Excel doesn't allow a single cell to be merged
      raise "Can't merge single cell" if row_first == row_last && col_first == col_last

      # Swap last row/col with first row/col as necessary
      row_first, row_last = row_last, row_first if row_first > row_last
      col_first, col_last = col_last, col_first if col_first > col_last

      # Check that column number is valid and store the max value
      return if check_dimensions(row_last, col_last) != 0

      # Store the merge range.
      @merge.push([row_first, col_first, row_last, col_last])

      # Write the first cell
      if type == 'string'
        write_string(row_first, col_first, *args)
      elsif type == 'number'
        write_number( row_first, col_first, *args)
      elsif type == 'blank'
        write_blank( row_first, col_first, *args)
      elsif type == 'date_time'
        write_date_time( row_first, col_first, *args)
      elsif type == 'rich_string'
        write_rich_string( row_first, col_first, *args)
      elsif type == 'url'
        write_url( row_first, col_first, *args)
      elsif type == 'formula'
        write_formula( row_first, col_first, *args)
      elsif type == 'array_formula'
        write_formula_array( row_first, col_first, *args)
      else
        raise "Unknown type '#{type}'"
      end

      # Pad out the rest of the area with formatted blank cells.
      (row_first .. row_last).each do |row|
        (col_first .. col_last).each do |col|
          next if row == row_first && col == col_first
          write_blank( row, col, format )
        end
      end
    end

    #
    # conditional_formatting($row, $col, {...})
    #
    # This method handles the interface to Excel conditional formatting.
    #
    # We allow the format to be called on one cell or a range of cells. The
    # hashref contains the formatting parameters and must be the last param:
    #    conditional_formatting($row, $col, {...})
    #    conditional_formatting($first_row, $first_col, $last_row, $last_col, {...})
    #
    # Returns  0 : normal termination
    #         -1 : insufficient number of arguments
    #         -2 : row or column out of range
    #         -3 : incorrect parameter.
    #
    def conditional_formatting(*args)
      # Check for a cell reference in A1 notation and substitute row and column
      args = row_col_notation(args)

      # Check for a valid number of args.
      return -1 unless args.size == 3 || args.size == 5

      # The final hash contains the validation parameters.
      param = args.pop

      # Make the last row/col the same as the first if not defined.
      row1, col1, row2, col2 = args
      row2, col2 = row1, col1 unless row2

      # Check that row and col are valid without storing the values.
      return -2 if check_dimensions(row1, col1, 1, 1) != 0
      return -2 if check_dimensions(row2, col2, 1, 1) != 0

      # List of valid input parameters.
      valid_parameter = {
          :type     => 1,
          :format   => 1,
          :criteria => 1,
          :value    => 1,
          :minimum  => 1,
          :maximum  => 1
      }
      # Check for valid input parameters.
      param.keys.each {|key| return -3 unless valid_parameter.keys.include?(key) }
      return -3 unless param.has_key?(:type)

      # List of  valid validation types.
      valid_type = { 'cell' => 'cellIs' }

      # Check for valid validation types.
      if valid_type.has_key?(param[:type].downcase)
        param[:type] = valid_type[param[:type].downcase]
      else
        return -3
      end

      # 'criteria' is a required parameter.
      return -3 unless param.has_key?(:criteria)

      # List of valid criteria types.
      criteria_type = {
          'between'                  => 'between',
          'not between'              => 'notBetween',
          'equal to'                 => 'equal',
          '='                        => 'equal',
          '=='                       => 'equal',
          'not equal to'             => 'notEqual',
          '!='                       => 'notEqual',
          '<>'                       => 'notEqual',
          'greater than'             => 'greaterThan',
          '>'                        => 'greaterThan',
          'less than'                => 'lessThan',
          '<'                        => 'lessThan',
          'greater than or equal to' => 'greaterThanOrEqual',
          '>='                       => 'greaterThanOrEqual',
          'less than or equal to'    => 'lessThanOrEqual',
          '<='                       => 'lessThanOrEqual'
      }

      # Check for valid criteria types.
      if criteria_type.has_key?(param[:criteria].downcase)
        param[:criteria] = criteria_type[param[:criteria].downcase]
      else
        return -3
      end

      # 'Between' and 'Not between' criteria require 2 values.
      if param[:criteria] == 'between' || param[:criteria] == 'notBetween'
        return -3 unless param.has_key?(:minimum)
        return -3 unless param.has_key?(:maximum)
      else
        param[:minimum] = nil
        param[:maximum] = nil
      end

      # Convert date/times value if required.
      if param[:type] == 'date' || param[:type] == 'time'
        if param[:value] =~ /T/
          date_time = convert_date_time(param[:value])
          if date_time
            param[:value] = date_time
          else
            return -3
          end
        end
        if param[:maximum] && param[:maximum] =~ /T/
          date_time = convert_date_time(param[:maximum])
          if date_time
            param[:maximum] = date_time
          else
            return -3
          end
        end
      end

      # Set the formatting range.
      range = ''

      # Swap last row/col for first row/col as necessary
      row1, row2 = row2, row1 if row1 > row2
      col1, col2 = col2, col1 if col1 > col2

      # If the first and last cell are the same write a single cell.
      if row1 == row2 && col1 == col2
        range = xl_rowcol_to_cell(row1, col1)
      else
        range = xl_range(row1, row2, col1, col2)
      end

      # Get the dxf format index.
      if param[:format]
        param[:format] = param[:format].get_dxf_index
      end

      # Set the priority based on the order of adding.
      param[:priority] = @dxf_priority
      @dxf_priority += 1

      # Store the validation information until we close the worksheet.
      @cond_formats[range] ||= []
      @cond_formats[range] << param
    end

    def data_validation(*args)
      # Check for a cell reference in A1 notation and substitute row and column
      args = row_col_notation(args)

      # Check for a valid number of args.
      return -1 if args.size != 5 && args.size != 3

      # The final hashref contains the validation parameters.
      param = args.pop

      # Make the last row/col the same as the first if not defined.
      row1, col1, row2, col2 = args
      unless row2
        row2 = row1
        col2 = col1
      end

      # Check that row and col are valid without storing the values.
      return -2 if check_dimensions(row1, col1, 1, 1) != 0
      return -2 if check_dimensions(row2, col2, 1, 1) != 0

      # Check that the last parameter is a hash list.
      unless param.respond_to?(:to_hash)
        #           carp "Last parameter '$param' in data_validation() must be a hash ref"
        return -3
      end

      # List of valid input parameters.
      valid_parameter = {
        :validate          => 1,
        :criteria          => 1,
        :value             => 1,
        :source            => 1,
        :minimum           => 1,
        :maximum           => 1,
        :ignore_blank      => 1,
        :dropdown          => 1,
        :show_input        => 1,
        :input_title       => 1,
        :input_message     => 1,
        :show_error        => 1,
        :error_title       => 1,
        :error_message     => 1,
        :error_type        => 1,
        :other_cells       => 1
      }

      # Check for valid input parameters.
      param.each_key do |param_key|
        unless valid_parameter.has_key?(param_key)
          #               carp "Unknown parameter '$param_key' in data_validation()"
          return -3
        end
      end

      # Map alternative parameter names 'source' or 'minimum' to 'value'.
      param[:value] = param[:source]  if param[:source]
      param[:value] = param[:minimum] if param[:minimum]

      # 'validate' is a required parameter.
      unless param.has_key?(:validate)
        #           carp "Parameter 'validate' is required in data_validation()"
        return -3
      end

      # List of  valid validation types.
      valid_type = {
        'any'             => 'none',
        'any value'       => 'none',
        'whole number'    => 'whole',
        'whole'           => 'whole',
        'integer'         => 'whole',
        'decimal'         => 'decimal',
        'list'            => 'list',
        'date'            => 'date',
        'time'            => 'time',
        'text length'     => 'textLength',
        'length'          => 'textLength',
        'custom'          => 'custom'
      }

      # Check for valid validation types.
      unless valid_type.has_key?(param[:validate].downcase)
        #           carp "Unknown validation type '$param->{validate}' for parameter " .
        #                "'validate' in data_validation()"
        return -3
      else
        param[:validate] = valid_type[param[:validate].downcase]
      end

      # No action is required for validation type 'any'.
      # TODO: we should perhaps store 'any' for message only validations.
      return 0 if param[:validate] == 0

      # The list and custom validations don't have a criteria so we use a default
      # of 'between'.
      if param[:validate] == 'list' || param[:validate] == 'custom'
        param[:criteria]  = 'between'
        param[:maximum]   = nil
      end

      # 'criteria' is a required parameter.
      unless param.has_key?(:criteria)
        #           carp "Parameter 'criteria' is required in data_validation()"
        return -3
      end

      # List of valid criteria types.
      criteria_type = {
        'between'                     => 'between',
        'not between'                 => 'notBetween',
        'equal to'                    => 'equal',
        '='                           => 'equal',
        '=='                          => 'equal',
        'not equal to'                => 'notEqual',
        '!='                          => 'notEqual',
        '<>'                          => 'notEqual',
        'greater than'                => 'greaterThan',
        '>'                           => 'greaterThan',
        'less than'                   => 'lessThan',
        '<'                           => 'lessThan',
        'greater than or equal to'    => 'greaterThanOrEqual',
        '>='                          => 'greaterThanOrEqual',
        'less than or equal to'       => 'lessThanOrEqual',
        '<='                          => 'lessThanOrEqual'
      }

      # Check for valid criteria types.
      unless criteria_type.has_key?(param[:criteria].downcase)
        #           carp "Unknown criteria type '$param->{criteria}' for parameter " .
        #                "'criteria' in data_validation()"
        return -3
      else
        param[:criteria] = criteria_type[param[:criteria].downcase]
      end

      # 'Between' and 'Not between' criteria require 2 values.
      if param[:criteria] == 'between' || param[:criteria] == 'notBetween'
        unless param.has_key?(:maximum)
          #               carp "Parameter 'maximum' is required in data_validation() " .
          #                    "when using 'between' or 'not between' criteria"
          return -3
        end
      else
        param[:maximum] = nil
      end

      # List of valid error dialog types.
      error_type = {
        'stop'        => 0,
        'warning'     => 1,
        'information' => 2
      }

      # Check for valid error dialog types.
      if not param.has_key?(:error_type)
        param[:error_type] = 0
      elsif not error_type.has_key?(param[:error_type].downcase)
        #           carp "Unknown criteria type '$param->{error_type}' for parameter " .
        #                "'error_type' in data_validation()"
        return -3
      else
        param[:error_type] = error_type[param[:error_type].downcase]
      end

      # Convert date/times value if required.
      if param[:validate] == 'date' || param[:validate] == 'time'
        if param[:value] =~ /T/
          date_time = convert_date_time(param[:value])
          unless date_time
            #                   carp "Invalid date/time value '$param->{value}' " .
            #                        "in data_validation()"
            return -3
          else
            param[:value] = date_time
          end
        end
        if param[:maximum] && param[:maximum] =~ /T/
          date_time = convert_date_time(param[:maximum])

          unless date_time
            #                   carp "Invalid date/time value '$param->{maximum}' " .
            #                        "in data_validation()"
            return -3
          else
            param[:maximum] = date_time
          end
        end
      end

      # Set some defaults if they haven't been defined by the user.
      param[:ignore_blank]  = 1 unless param[:ignore_blank]
      param[:dropdown]      = 1 unless param[:dropdown]
      param[:show_input]    = 1 unless param[:show_input]
      param[:show_error]    = 1 unless param[:show_error]

      # These are the cells to which the validation is applied.
      param[:cells] = [[row1, col1, row2, col2]]

      # A (for now) undocumented parameter to pass additional cell ranges.
      if param.has_key?(:other_cells)
        param[:other_cells].each { |cells| param[:cells] << cells }
      end

      # Store the validation information until we close the worksheet.
      @validations.push(param)
    end

    #
    # Set the option to hide gridlines on the screen and the printed page.
    #
    # This was mainly useful for Excel 5 where printed gridlines were on by
    # default.
    #
    def hide_gridlines(option = true)
      if option == true
        @print_gridlines  = false
        @screen_gridlines = true
      elsif !option
        @print_gridlines       = true    # 1 = display, 0 = hide
        @screen_gridlines      = true
        @print_options_changed = true
      else
        @print_gridlines  = false
        @screen_gridlines = false
      end
    end

    #
    # autofilter($first_row, $first_col, $last_row, $last_col)
    #
    # Set the autofilter area in the worksheet.
    #
    def autofilter(*args)
      # Check for a cell reference in A1 notation and substitute row and column
      args = row_col_notation(args)

      return if args.size != 4    # Require 4 parameters

      row1, col1, row2, col2 = args

      # Reverse max and min values if necessary.
      row1, row2 = row2, row1 if row2 < row1
      col1, col2 = col2, col1 if col2 < col1

      # Build up the print area range "Sheet1!$A$1:$C$13".
      area = convert_name_area(row1, col1, row2, col2)
      ref = xl_range(row1, row2, col1, col2)

      @autofilter_area = area
      @autofilter_ref  = ref
      @filter_range    = [col1, col2]
    end

    #
    # Set the column filter criteria.
    #
    def filter_column(col, expression)
      raise "Must call autofilter before filter_column" unless @autofilter_area

      # Check for a column reference in A1 notation and substitute.
      if col =~ /^\D/
        col_letter = col

        # Convert col ref to a cell ref and then to a col number.
        dummy, col = substitute_cellref("#{col}1")
        raise "Invalid column '#{col_letter}'" if col >= @xls_colmax
      end

      col_first, col_last = @filter_range

      # Reject column if it is outside filter range.
      if col < col_first or col > col_last
        raise "Column '#{col}' outside autofilter column range (#{col_first} .. #{col_last})"
      end

      tokens = extract_filter_tokens(expression)

      unless tokens.size == 3 || tokens.size == 7
        raise "Incorrect number of tokens in expression '#{expression}'"
      end

      tokens = parse_filter_expression(expression, tokens)

      # Excel handles single or double custom filters as default filters. We need
      # to check for them and handle them accordingly.
      if tokens.size == 2 && tokens[0] == 2
        # Single equality.
        filter_column_list(col, tokens[1])
      elsif tokens.size == 5 && tokens[0] == 2 && tokens[2] == 1 && tokens[3] == 2
        # Double equality with "or" operator.
        filter_column_list(col, tokens[1], tokens[4])
      else
        # Non default custom filter.
        @filter_cols[col] = Array.new(tokens)
        @filter_type[col] = 0
      end

      @filter_on = 1
    end

    #
    # Set the column filter criteria in Excel 2007 list style.
    #
    def filter_column_list(col, *tokens)
      tokens.flatten!
      raise "Must call autofilter before filter_column_list" unless @autofilter_area
      raise "Incorrect number of arguments to filter_column_list" if tokens.empty?

      # Check for a column reference in A1 notation and substitute.
      if col =~ /^\D/
        col_letter = col

        # Convert col ref to a cell ref and then to a col number.
        dummy, col = substitute_cellref("#{col}1")
        raise "Invalid column '#{col_letter}'" if col >= @xls_colmax
      end

      col_first, col_last = @filter_range

      # Reject column if it is outside filter range.
      if col < col_first || col > col_last
        raise "Column '#{col}' outside autofilter column range (#{col_first} .. #{col_last})"
      end

      @filter_cols[col] = tokens
      @filter_type[col] = 1           # Default style.
      @filter_on        = 1
    end

    #
    # Store the horizontal page breaks on a worksheet.
    #
    def set_h_pagebreaks(*args)
      @hbreaks += args
    end

    #
    # Store the vertical page breaks on a worksheet.
    #
    def set_v_pagebreaks(*args)
      @vbreaks += args
    end

    #
    # Make any comments in the worksheet visible.
    #
    def show_comments(visible = true)
      @comments_visible = visible
    end

    def has_comments?
      !!@has_comments
    end

    def is_chartsheet?
      !!@is_chartsheet
    end

    #
    # Turn the HoH that stores the comments into an array for easier handling
    # and set the external links.
    #
    def prepare_comments(vml_data_id, vml_shape_id, comment_id)
      comments = []

      # We sort the comments by row and column but that isn't strictly required.
      @comments.keys.sort.each do |row|
        @comments[row].keys.sort.each do |col|
          # Set comment visibility if required and not already user defined.
          @comments[row][col][4] ||= 1 if @comments_visible

          # Set comment author if not already user defined.
          @comments[row][col][3] ||= @comments_author
          comments << @comments[row][col]
        end
      end

      @comments_array = comments

      @external_comment_links <<
        ['/vmlDrawing', "../drawings/vmlDrawing#{comment_id}.vml"] <<
        ['/comments',   "../comments#{comment_id}.xml"]

      count         = comments.size
      start_data_id = vml_data_id

      # The VML o:idmap data id contains a comma separated range when there is
      # more than one 1024 block of comments, like this: data="1,2".
      ( 1 .. ( count / 1024 ) ).each do |i|
        vml_data_id = "vml_data_id,#{start_data_id + i}"
      end

      @vml_data_id  = vml_data_id
      @vml_shape_id = vml_shape_id

      count
    end

    #
    # Set up chart/drawings.
    #
    def prepare_chart(index, chart_id, drawing_id)
      drawing_type = 1

      row, col, chart, x_offset, y_offset, scale_x, scale_y  = @charts[index]
      scale_x ||= 0
      scale_y ||= 0

      width  = ( 0.5 + ( 480 * scale_x ) ).to_i
      height = ( 0.5 + ( 288 * scale_y ) ).to_i

      dimensions = position_object_emus(col, row, x_offset, y_offset, width, height)

      # Create a Drawing object to use with worksheet unless one already exists.
      if !drawing?
        drawing = Drawing.new
        drawing.add_drawing_object(drawing_type, dimensions)
        drawing.embedded = 1

        @drawing = drawing

        @external_drawing_links << ['/drawing', "../drawings/drawing#{drawing_id}.xml" ]
      else
        @drawing.add_drawing_object(drawing_type, dimensions)
      end
      @drawing_links << ['/chart', "../charts/chart#{chart_id}.xml"]
    end

    #
    # Returns a range of data from the worksheet _table to be used in chart
    # cached data. Strings are returned as SST ids and decoded in the workbook.
    # Return undefs for data that doesn't exist since Excel can chart series
    # with data missing.
    #
    def get_range_data(row_start, col_start, row_end, col_end)
      # TODO. Check for worksheet limits.

      # Iterate through the table data.
      data = []
      (row_start .. row_end).each do |row_num|
        # Store undef if row doesn't exist.
        if !@table[row_num]
          data << nil
          next
        end

        (col_start .. col_end).each do |col_num|
          if cell = @table[row_num][col_num]
            type  = cell[0]
            token = cell[1]

            data << case type
            when 'n'
              # Store a number.
              token
            when 's'
              # Store a string.
              {:sst_id => token}
            when 'f'
              # Store a formula.
              cell[3] || 0
            when 'a'
              # Store an array formula.
              cell[4] || 0
            when 'l'
              # Store the string part a hyperlink.
              {:sst_id => token}
            when 'b'
              # Store a empty cell.
              ''
            end
          else
            # Store undef if col doesn't exist.
            data << nil
          end
        end
      end

      return data
    end

    private

    #
    # Extract the tokens from the filter expression. The tokens are mainly non-
    # whitespace groups. The only tricky part is to extract string tokens that
    # contain whitespace and/or quoted double quotes (Excel's escaped quotes).
    #
    # Examples: 'x <  2000'
    #           'x >  2000 and x <  5000'
    #           'x = "foo"'
    #           'x = "foo bar"'
    #           'x = "foo "" bar"'
    #
    def extract_filter_tokens(expression = nil)
      return [] unless expression

      tokens = []
      str = expression
      while str =~ /"(?:[^"]|"")*"|\S+/
        tokens << $&
        str = $~.post_match
      end

      # Remove leading and trailing quotes and unescape other quotes
      tokens.map! do |token|
        token.sub!(/^"/, '')
        token.sub!(/"$/, '')
        token.gsub!(/""/, '"')

        # if token is number, convert to numeric.
        if token =~ /^([+-]?)(?=\d|\.\d)\d*(\.\d*)?([Ee]([+-]?\d+))?$/
          token.to_f == token.to_i ? token.to_i : token.to_f
        else
          token
        end
      end

      tokens
    end

    #
    # Converts the tokens of a possibly conditional expression into 1 or 2
    # sub expressions for further parsing.
    #
    # Examples:
    #          ('x', '==', 2000) -> exp1
    #          ('x', '>',  2000, 'and', 'x', '<', 5000) -> exp1 and exp2
    #
    def parse_filter_expression(expression, tokens)
      # The number of tokens will be either 3 (for 1 expression)
      # or 7 (for 2  expressions).
      #
      if (tokens.size == 7)
        conditional = tokens[3]
        if conditional =~ /^(and|&&)$/
          conditional = 0
        elsif conditional =~ /^(or|\|\|)$/
          conditional = 1
        else
          raise "Token '#{conditional}' is not a valid conditional " +
          "in filter expression '#{expression}'"
        end
        expression_1 = parse_filter_tokens(expression, tokens[0..2])
        expression_2 = parse_filter_tokens(expression, tokens[4..6])
        [expression_1, conditional, expression_2].flatten
      else
        parse_filter_tokens(expression, tokens)
      end
    end

    #
    # Parse the 3 tokens of a filter expression and return the operator and token.
    #
    def parse_filter_tokens(expression, tokens)     #:nodoc:
      operators = {
        '==' => 2,
        '='  => 2,
        '=~' => 2,
        'eq' => 2,

        '!=' => 5,
        '!~' => 5,
        'ne' => 5,
        '<>' => 5,

        '<'  => 1,
        '<=' => 3,
        '>'  => 4,
        '>=' => 6,
      }

      operator = operators[tokens[1]]
      token    = tokens[2]

      # Special handling of "Top" filter expressions.
      if tokens[0] =~ /^top|bottom$/i
        value = tokens[1]
        if (value =~ /\D/ or value.to_i < 1 or value.to_i > 500)
          raise "The value '#{value}' in expression '#{expression}' " +
          "must be in the range 1 to 500"
        end
        token.downcase!
        if (token != 'items' and token != '%')
          raise "The type '#{token}' in expression '#{expression}' " +
          "must be either 'items' or '%'"
        end

        if (tokens[0] =~ /^top$/i)
          operator = 30
        else
          operator = 32
        end

        if (tokens[2] == '%')
          operator += 1
        end

        token    = value
      end

      if (not operator and tokens[0])
        raise "Token '#{tokens[1]}' is not a valid operator " +
        "in filter expression '#{expression}'"
      end

      # Special handling for Blanks/NonBlanks.
      if (token =~ /^blanks|nonblanks$/i)
        # Only allow Equals or NotEqual in this context.
        if (operator != 2 and operator != 5)
          raise "The operator '#{tokens[1]}' in expression '#{expression}' " +
          "is not valid in relation to Blanks/NonBlanks'"
        end

        token.downcase!

        # The operator should always be 2 (=) to flag a "simple" equality in
        # the binary record. Therefore we convert <> to =.
        if token == 'blanks'
          if operator == 5
            token = ' '
          end
        else
          if operator == 5
            operator = 2
            token    = 'blanks'
          else
            operator = 5
            token    = ' '
          end
        end
      end

      # if the string token contains an Excel match character then change the
      # operator type to indicate a non "simple" equality.
      if (operator == 2 and token =~ /[*?]/)
        operator = 22
      end

      [operator, token]
    end

    #
    # Convert from an Excel internal colour index to a XML style #RRGGBB index
    # based on the default or user defined values in the Workbook palette.
    #
    def get_palette_color(index)
      # Adjust the colour index.
      index -= 8

      # Palette is passed in from the Workbook class.
      rgb = @workbook.palette[index]

      # TODO Add the alpha part to the RGB.
      sprintf("FF%02X%02X%02X", *rgb[0, 3])
    end

    #
    # Substitute an Excel cell reference in A1 notation for  zero based row and
    # column values in an argument list.
    #
    # Ex: ("A4", "Hello") is converted to (3, 0, "Hello").
    #
    def substitute_cellref(cell, *args)       #:nodoc:
      return [*args] if cell.respond_to?(:coerce) # Numeric

      cell.upcase!

      case cell
      # Convert a column range: 'A:A' or 'B:G'.
      # A range such as A:A is equivalent to A1:65536, so add rows as required
      when /\$?([A-Z]{1,3}):\$?([A-Z]{1,3})/
        row1, col1 =  cell_to_rowcol($1 + '1')
        row2, col2 =  cell_to_rowcol($2 + @xls_rowmax.to_s)
        return [row1, col1, row2, col2, *args]
      # Convert a cell range: 'A1:B7'
      when /\$?([A-Z]{1,3}\$?\d+):\$?([A-Z]{1,3}\$?\d+)/
        row1, col1 =  cell_to_rowcol($1)
        row2, col2 =  cell_to_rowcol($2)
        return [row1, col1, row2, col2, *args]
      # Convert a cell reference: 'A1' or 'AD2000'
      when /\$?([A-Z]{1,3}\$?\d+)/
        row1, col1 =  cell_to_rowcol($1)
        return [row1, col1, *args]
      else
        raise("Unknown cell reference #{cell}")
      end
    end

    #
    # Convert an Excel cell reference in A1 notation to a zero based row and column
    # reference converts C1 to (0, 2).
    #
    # Returns: row, column
    #
    def cell_to_rowcol(cell)       #:nodoc:
      cell =~ /(\$?)([A-Z]{1,3})(\$?)(\d+)/
      col_abs = $1 == '' ? 0 : 1
      col     = $2
      row_abs = $3 == '' ? 0 : 1
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

      [row, col, row_abs, col_abs]
    end

    #
    # This is an internal method that is used to filter elements of the array of
    # pagebreaks used in the _store_hbreak() and _store_vbreak() methods. It:
    #   1. Removes duplicate entries from the list.
    #   2. Sorts the list.
    #   3. Removes 0 from the list if present.
    #
    def sort_pagebreaks(*args)
      return [] if args.empty?

      breaks = args.uniq.sort
      breaks.delete(0)

      # The Excel 2007 specification says that the maximum number of page breaks
      # is 1026. However, in practice it is actually 1023.
      max_num_breaks = 1023
      if breaks.size > max_num_breaks
        breaks[0, max_num_breaks]
      else
        breaks
      end
    end

    #
    # Calculate the vertices that define the position of a graphical object within
    # the worksheet in pixels.
    #
    #         +------------+------------+
    #         |     A      |      B     |
    #   +-----+------------+------------+
    #   |     |(x1,y1)     |            |
    #   |  1  |(A1)._______|______      |
    #   |     |    |              |     |
    #   |     |    |              |     |
    #   +-----+----|    BITMAP    |-----+
    #   |     |    |              |     |
    #   |  2  |    |______________.     |
    #   |     |            |        (B2)|
    #   |     |            |     (x2,y2)|
    #   +---- +------------+------------+
    #
    # Example of an object that covers some of the area from cell A1 to cell B2.
    #
    # Based on the width and height of the object we need to calculate 8 vars:
    #
    #     $col_start, $row_start, $col_end, $row_end, $x1, $y1, $x2, $y2.
    #
    # We also calculate the absolute x and y position of the top left vertex of
    # the object. This is required for images.
    #
    #    $x_abs, $y_abs
    #
    # The width and height of the cells that the object occupies can be variable
    # and have to be taken into account.
    #
    # The values of $col_start and $row_start are passed in from the calling
    # function. The values of $col_end and $row_end are calculated by subtracting
    # the width and height of the object from the width and height of the
    # underlying cells.
    #
    #    col_start    # Col containing upper left corner of object.
    #    x1           # Distance to left side of object.
    #    row_start    # Row containing top left corner of object.
    #    y1           # Distance to top of object.
    #    col_end      # Col containing lower right corner of object.
    #    x2           # Distance to right side of object.
    #    row_end      # Row containing bottom right corner of object.
    #    y2           # Distance to bottom of object.
    #    width        # Width of object frame.
    #    height       # Height of object frame.
    def position_object_pixels(col_start, row_start, x1, y1, width, height, is_drawing = false)
      x_abs = 0    # Absolute distance to left side of object.
      y_abs = 0    # Absolute distance to top  side of object.

      # Calculate the absolute x offset of the top-left vertex.
      if @col_size_changed
        (1 .. col_start).each {|col_id| x_abs += size_col(col_id) }
      else
        # Optimisation for when the column widths haven't changed.
        x_abs += 64 * col_start
      end

      x_abs += x1

      # Calculate the absolute y offset of the top-left vertex.
      # Store the column change to allow optimisations.
      if @row_size_changed
        (1 .. row_start).each {|row_id| y_abs += size_row(row_id) }
      else
        # Optimisation for when the row heights haven't changed.
        y_abs += 20 * row_start
      end

      y_abs += y1

      # Adjust start column for offsets that are greater than the col width.
      while x1 >= size_col(col_start)
        x1 -= size_col(col_start)
        col_start += 1
      end

      # Adjust start row for offsets that are greater than the row height.
      while y1 >= size_row(row_start)
        y1 -= size_row(row_start)
        row_start += 1
      end


      # Initialise end cell to the same as the start cell.
      col_end = col_start
      row_end = row_start

      width  += x1
      height += y1

      # Subtract the underlying cell widths to find the end cell of the object.
      while width >= size_col(col_end)
        width -= size_col(col_end)
        col_end += 1
      end

      # Subtract the underlying cell heights to find the end cell of the object.
      while height >= size_row(row_end)
        height -= size_row(row_end)
        row_end += 1
      end

      # The following is only required for positioning drawing/chart objects
      # and not comments. It is probably the result of a bug.
      if is_drawing
        col_end -= 1 if width == 0
        row_end -= 1 if height == 0
      end

      # The end vertices are whatever is left from the width and height.
      x2 = width
      y2 = height

      [col_start, row_start, x1, y1, col_end, row_end, x2, y2, x_abs, y_abs]
    end

    #
    # Calculate the vertices that define the position of a graphical object within
    # the worksheet in EMUs.
    #
    # The vertices are expressed as English Metric Units (EMUs). There are 12,700
    # EMUs per point. Therefore, 12,700 * 3 /4 = 9,525 EMUs per pixel.
    #
    def position_object_emus(col_start, row_start, x1, y1, width, height)
      is_drawing = true
      col_start, row_start, x1, y1, col_end, row_end, x2, y2, x_abs, y_abs =
        position_object_pixels( col_start, row_start, x1, y1, width, height, is_drawing)

      # Convert the pixel values to EMUs. See above.
      x1    *= 9_525
      y1    *= 9_525
      x2    *= 9_525
      y2    *= 9_525
      x_abs *= 9_525
      y_abs *= 9_525

      [col_start, row_start, x1, y1, col_end, row_end, x2, y2, x_abs, y_abs]
    end

    #
    # Convert the width of a cell from user's units to pixels. Excel rounds the
    # column width to the nearest pixel. If the width hasn't been set by the user
    # we use the default value. If the column is hidden it has a value of zero.
    #
    def size_col(col)
      max_digit_width = 7    # For Calabri 11.
      padding         = 5

      # Look up the cell value to see if it has been changed.
      if @col_sizes[col]
        width = @col_sizes[col]

        # Convert to pixels.
        if width == 0
          pixels = 0
        elsif width < 1
          pixels = (width * 12 + 0.5).to_i
        else
          pixels = (width * max_digit_width + 0.5).to_i + padding
        end
      else
        pixels = 64
      end
      pixels
    end

    #
    # Convert the height of a cell from user's units to pixels. If the height
    # hasn't been set by the user we use the default value. If the row is hidden
    # it has a value of zero.
    #
    def size_row(row)
      # Look up the cell value to see if it has been changed
      if @row_sizes[row]
        height = @row_sizes[row]

        if height == 0
          pixels = 0
        else
          pixels = (4 / 3.0 * height).to_i
        end
      else
        pixels = 20
      end
      pixels
    end

    #
    # Add a string to the shared string table, if it isn't already there, and
    # return the string index.
    #
    def get_shared_string_index(str)
      # Add the string to the shared string table.
      unless @workbook.str_table[str]
        @workbook.str_table[str] = @workbook.str_unique
        @workbook.str_unique += 1
      end

      @workbook.str_total += 1
      index = @workbook.str_table[str]
    end

    #
    # Set up image/drawings.
    #
    def prepare_image(index, image_id, drawing_id, width, height, name, image_type)
      drawing_type = 2
      drawing

      row, col, image, x_offset, y_offset, scale_x, scale_y = @images[index]

      width  *= scale_x
      height *= scale_y

      dimensions = position_object_emus(col, row, x_offset, y_offset, width, height)

      # Convert from pixels to emus.
      width  = int( 0.5 + ( width * 9_525 ) )
      height = int( 0.5 + ( height * 9_525 ) )

      # Create a Drawing object to use with worksheet unless one already exists.
      if !drawing?
        drawing = Drawing.new
        drawing.embedded = 1

        @drawing = drawing

        @external_drawing_links << ['/drawing', "../drawings/drawing#{drawing_id}.xml"]
      else
        drawing = @drawing
      end

      drawing.add_drawing_object(drawing_type, dimensions, width, height, name)

      @drawing_links << ['/image', "../media/image#{image_id}.#{image_type}"]
    end

    #
    # This method handles the additional optional parameters to write_comment() as
    # well as calculating the comment object position and vertices.
    #
    def comment_params(row, col, string, options = {})
      default_width  = 128
      default_height = 74

      params = {
        :author          => nil,
        :color           => 81,
        :start_cell      => nil,
        :start_col       => nil,
        :start_row       => nil,
        :visible         => nil,
        :width           => default_width,
        :height          => default_height,
        :x_offset        => nil,
        :x_scale         => 1,
        :y_offset        => nil,
        :y_scale         => 1
      }

      # Overwrite the defaults with any user supplied values. Incorrect or
      # misspelled parameters are silently ignored.
      params.update(options)

      # Ensure that a width and height have been set.
      params[:width]  ||= default_width
      params[:height] ||= default_height

      # Limit the string to the max number of chars.
      max_len = 32767

      string = string[0, max_len] if string.length > max_len

      # Set the comment background colour.
      color    = params[:color]
      color_id = Format.get_color(color)

      if color_id == 0
        params[:color] = '#ffffe1'
      else
        # Get the RGB color from the palette.
        rgb = @workbook.palette[color_id - 8]

        # Minor modification to allow comparison testing. Change RGB colors
        # from long format, ffcc00 to short format fc0 used by VML.
        rgb_color = sprintf("%02x%02x%02x", *rgb)

        if rgb_color =~ /^([0-9a-f])\1([0-9a-f])\2([0-9a-f])\3$/
          rgb_color = "#{$1}#{$2}#{$3}"
        end

        params[:color] = sprintf("#%s [%d]\n", rgb_color, color_id)
      end

      # Convert a cell reference to a row and column.
      if params[:start_cell]
        params[:start_row], params[:start_col] = substitute_cellref(params[:start_cell])
      end

      # Set the default start cell and offsets for the comment. These are
      # generally fixed in relation to the parent cell. However there are
      # some edge cases for cells at the, er, edges.
      #
      row_max = @xls_rowmax
      col_max = @xls_colmax

      params[:start_row] ||= case row
        when 0
          0
        when row_max - 3
          row_max - 7
        when row_max - 2
          row_max - 6
        when row_max - 1
          row_max - 5
        else
          row - 1
      end

      params[:y_offset] ||= case row
        when 0
          2
        when row_max - 3, row_max - 2
          16
        when row_max - 1
          14
        else
          10
      end

      params[:start_col] ||= case col
        when col_max - 3
          col_max - 6
        when col_max - 2
          col_max - 5
        when col_max - 1
          col_max - 4
        else
          col + 1
      end

      params[:x_offset] ||= case col
        when col_max - 3, col_max - 2, col_max - 1
          49
        else
          15
      end

      # Scale the size of the comment box if required.
      params[:width] = params[:width] * params[:x_scale] if params[:x_scale]

      params[:height] = params[:height] * params[:y_scale] if params[:y_scale]

      # Round the dimensions to the nearest pixel.
      params[:width]  = ( 0.5 + params[:width] ).to_i
      params[:height] = ( 0.5 + params[:height] ).to_i

      # Calculate the positions of comment object.
      vertices = position_object_pixels(
        params[:start_col], params[:start_row], params[:x_offset],
        params[:y_offset],  params[:width],     params[:height]
        )

      # Add the width and height for VML.
      vertices << [params[:width], params[:height]]

      return [
        row,
        col,
        string,

        params[:author],
        params[:visible],
        params[:color],

        vertices
      ]
    end

    #
    # Based on the algorithm provided by Daniel Rentz of OpenOffice.
    #
    def encode_password(password)
      i = 0
      chars = password.split(//)
      count = chars.size

      chars.collect! do |char|
        i += 1
        char     = char.ord << i
        low_15   = char & 0x7fff
        high_15  = char & 0x7fff << 15
        high_15  = high_15 >> 15
        char     = low_15 | high_15
      end

      encoded_password  = 0x0000
      chars.each { |c| encoded_password ^= c }
      encoded_password ^= count
      encoded_password ^= 0xCE4B
    end

    #
    # Write the <worksheet> element. This is the root element of Worksheet.
    #
    def write_worksheet
        schema                 = 'http://schemas.openxmlformats.org/'
        attributes = [
          'xmlns',    schema + 'spreadsheetml/2006/main',
          'xmlns:r',  schema + 'officeDocument/2006/relationships'
        ]
        @writer.start_tag('worksheet', attributes)
    end

    #
    # Write the <sheetPr> element for Sheet level properties.
    #
    def write_sheet_pr
      return if !fit_page? && !filter_on? && !tab_color?
      attributes = []
      (attributes << 'filterMode' << 1) if filter_on?

      if fit_page? || tab_color?
        @writer.start_tag('sheetPr', attributes)
        write_tab_color
        write_page_set_up_pr
        @writer.end_tag('sheetPr')
      else
        @writer.empty_tag('sheetPr', attributes)
      end
    end

    #
    # Write the <pageSetUpPr> element.
    #
    def write_page_set_up_pr
      return unless fit_page?

      attributes = ['fitToPage', 1]
      @writer.empty_tag('pageSetUpPr', attributes)
    end

    # Write the <dimension> element. This specifies the range of cells in the
    # worksheet. As a special case, empty spreadsheets use 'A1' as a range.
    #
    def write_dimension
      if !@dim_rowmin && !@dim_colmin
        # If the min dims are undefined then no dimensions have been set
        # and we use the default 'A1'.
        ref = 'A1'
      elsif !@dim_rowmin && @dim_colmin
        # If the row dims aren't set but the column dims are then they
        # have been changed via set_column().
        if @dim_colmin == @dim_colmax
          # The dimensions are a single cell and not a range.
          ref = xl_rowcol_to_cell(0, @dim_colmin)
        else
          # The dimensions are a cell range.
          cell_1 = xl_rowcol_to_cell(0, @dim_colmin)
          cell_2 = xl_rowcol_to_cell(0, @dim_colmax)
          ref = cell_1 + ':' + cell_2
        end
      elsif @dim_rowmin == @dim_rowmax && @dim_colmin == @dim_colmax
        # The dimensions are a single cell and not a range.
        ref = xl_rowcol_to_cell(@dim_rowmin, @dim_colmin)
      else
        # The dimensions are a cell range.
        cell_1 = xl_rowcol_to_cell(@dim_rowmin, @dim_colmin)
        cell_2 = xl_rowcol_to_cell(@dim_rowmax, @dim_colmax)
        ref = cell_1 + ':' + cell_2
      end
      attributes = ['ref', ref]
      @writer.empty_tag('dimension', attributes)
    end
    #
    # Write the <sheetViews> element.
    #
    def write_sheet_views
      @writer.start_tag('sheetViews', [])
      write_sheet_view
      @writer.end_tag('sheetViews')
    end

    def write_sheet_view
      attributes = []
      # Hide screen gridlines if required
      attributes << 'showGridLines' << 0 unless screen_gridlines?

      # Hide zeroes in cells.
      attributes << 'showZeros' << 0 unless show_zeros?

      # Display worksheet right to left for Hebrew, Arabic and others.
      attributes << 'rightToLeft' << 1 if @right_to_left

      # Show that the sheet tab is selected.
      attributes << 'tabSelected' << 1 if @selected

      # Set the page view/layout mode if required.
      # TODO. Add pageBreakPreview mode when requested.
      (attributes << 'view' << 'pageLayout') if page_view?

      # Set the zoom level.
      if @zoom != 100
        (attributes << 'zoomScale' << @zoom) unless page_view?
        (attributes << 'zoomScaleNormal' << @zoom) if zoom_scale_normal?
      end

      attributes << 'workbookViewId' << 0

      if @panes.empty? && @selections.empty?
        @writer.empty_tag('sheetView', attributes)
      else
        @writer.start_tag('sheetView', attributes)
        write_panes
        write_selections
        @writer.end_tag('sheetView')
      end
    end

    #
    # Write the <selection> elements.
    #
    def write_selections
      @selections.each { |selection| write_selection(*selection) }
    end

    #
    # Write the <selection> element.
    #
    def write_selection(pane, active_cell, sqref)
      attributes  = []
      (attributes << 'pane' << pane) if pane
      (attributes << 'activeCell' << active_cell) if active_cell
      (attributes << 'sqref' << sqref) if sqref

      @writer.empty_tag('selection', attributes)
    end

    #
    # Write the <sheetFormatPr> element.
    #
    def write_sheet_format_pr
      base_col_width     = 10
      default_row_height = 15

      attributes = ['defaultRowHeight', default_row_height]
      attributes << 'outlineLevelRow' << @outline_row_level if @outline_row_level > 0
      attributes << 'outlineLevelCol' << @outline_col_level if @outline_col_level > 0
      @writer.empty_tag('sheetFormatPr', attributes)
    end

    #
    # Write the <cols> element and <col> sub elements.
    #
    def write_cols
      # Exit unless some column have been formatted.
      return if @colinfo.empty?

      @writer.start_tag('cols')
      @colinfo.each {|col_info| write_col_info(*col_info) }

      @writer.end_tag('cols')
    end

    #
    # Write the <col> element.
    #
    def write_col_info(*args)
      min = args[0] || 0
      max = args[1] || 0
      width = args[2]
      format = args[3]
      hidden = args[4] || 0
      level = args[5] || 0
      collapsed = args[6] || 0
#          min          = $_[0] // 0    # First formatted column.
#          max          = $_[1] // 0    # Last formatted column.
#          width        = $_[2]         # Col width in user units.
#          format       = $_[3]         # Format index.
#          hidden       = $_[4] // 0    # Hidden flag.
#          level        = $_[5] // 0    # Outline level.
#          collapsed    = $_[6] // 0    # Outline level.
      custom_width = 1
      xf_index = 0
      xf_index = format.get_xf_index if format.respond_to?(:get_xf_index)

      # Set the Excel default col width.
      if width.nil?
        if hidden == 0
          width        = 8.43
          custom_width = 0
        else
          width = 0
        end
      else
        # Width is defined but same as default.
        custom_width = 0 if width == 8.43
      end

      # Convert column width from user units to character width.
      max_digit_width = 7.0    # For Calabri 11.
      padding         = 5.0
      if width && width > 0
        width = ((width * max_digit_width + padding ) / max_digit_width * 256).to_i/256.0
      end
      attributes = [
          'min',   min + 1,
          'max',   max + 1,
          'width', width
      ]

      (attributes << 'style' << xf_index) if xf_index != 0
      (attributes << 'hidden' << 1)       if hidden != 0
      (attributes << 'customWidth' << 1)  if custom_width != 0
      (attributes << 'outlineLevel' << level) if level != 0
      (attributes << 'collapsed'    << 1) if collapsed != 0

      @writer.empty_tag('col', attributes)
    end

    #
    # Write the <sheetData> element.
    #
    def write_sheet_data
      if !@dim_rowmin
        # If the dimensions aren't defined then there is no data to write.
        @writer.empty_tag('sheetData')
      else
        @writer.start_tag('sheetData')
        write_rows
        @writer.end_tag('sheetData')
      end
    end

    #
    # Write out the worksheet data as a series of rows and cells.
    #
    def write_rows
      calculate_spans

      (@dim_rowmin .. @dim_rowmax).each do |row_num|
        # Skip row if it doesn't contain row formatting or cell data.
        next if !@set_rows[row_num] && !@table[row_num] && !@comments[row_num]

        span_index = row_num / 16
        span       = @row_spans[span_index]

        # Write the cells if the row contains data.
        if @table[row_num]
          if !@set_rows[row_num]
            write_row_element(row_num, span )
          else
            write_row_element(row_num, span, *(@set_rows[row_num]))
          end

          (@dim_colmin .. @dim_colmax).each do |col_num|
            col_ref = @table[row_num][col_num]
            write_cell(row_num, col_num, col_ref) if col_ref
          end

          @writer.end_tag('row')
        else
          # Row attributes only.
          write_empty_row(row_num, nil, *(@set_rows[row_num]))
        end
      end
    end

    #
    # Write out the worksheet data as a single row with cells. This method is
    # used when memory optimisation is on. A single row is written and the data
    # table is reset. That way only one row of data is kept in memory at any one
    # time. We don't write span data in the optimised case since it is optional.
    #
    def write_single_row(current_row = 0)
      row_num     = @previous_row

      # Set the new previous row as the current row.
      @previous_row = current_row

      # Skip row if it doesn't contain row formatting, cell data or a comment.
      return if !@set_rows[row_num] && !@table[row_num] && !@comments[row_num]

      # Write the cells if the row contains data.
      row_ref = @table[row_num]
      if row_ref
        if !@set_rows[row_num]
          write_row(row_num)
        else
          write_row(row_num, nil, @set_rows[row_num])
        end

        (@dim_colmin .. @dim_colmax).each do |col_num|
          col_ref = @table[row_num][col_num]
          write_cell(row_num, col_num, col_ref) if col_ref
        end
        @writer.end_tag('row')
      else
        # Row attributes or comments only.
        write_empty_row(row_num, nil, @set_rows[row_num])
      end

      # Reset table.
      @table = []
    end

    #
    # Write the <row> element.
    #
    def write_row_element(r, spans = nil, height = 15, format = nil, hidden = false, level = 0, collapsed = false, empty_row = false)
      height    ||= 15
      hidden    ||= 0
      level     ||= 0
      collapsed ||= 0
      empty_row ||= 0
      xf_index = 0

      attributes = ['r',  r + 1]

      xf_index = format.get_xf_index if format

      (attributes << 'spans'        << spans ) if spans
      (attributes << 's'            << xf_index) if xf_index != 0
      (attributes << 'customFormat' << 1     ) if format
      (attributes << 'ht'           << height) if height != 15
      (attributes << 'hidden'       << 1     ) if !!hidden && hidden != 0
      (attributes << 'customHeight' << 1     ) if height != 15
      (attributes << 'outlineLevel' << level ) if !!level && level != 0
      (attributes << 'collapsed'    << 1     ) if !!collapsed && collapsed != 0

      if empty_row && empty_row != 0
        @writer.empty_tag('row', attributes)
      else
        @writer.start_tag('row', attributes)
      end
    end

    #
    # Write and empty <row> element, i.e., attributes only, no cell data.
    #
    def write_empty_row(*args)
        new_args = args.dup
        new_args[7] = 1
        write_row_element(*new_args)
    end

    #
    # Write the <cell> element. This is the innermost loop so efficiency is
    # important where possible. The basic methodology is that the data of every
    # cell type is passed in as follows:
    #
    #      [ $row, $col, $aref]
    #
    # The aref, called $cell below, contains the following structure in all types:
    #
    #     [ $type, $token, $xf, @args ]
    #
    # Where $type:  represents the cell type, such as string, number, formula, etc.
    #       $token: is the actual data for the string, number, formula, etc.
    #       $xf:    is the XF format object index.
    #       @args:  additional args relevant to the specific data type.
    #
    def write_cell(row, col, cell)
      type, token, xf = cell

      xf_index = 0
      xf_index = xf.get_xf_index if xf.respond_to?(:get_xf_index)

      range = xl_rowcol_to_cell( row, col )
      attributes = ['r', range]

      # Add the cell format index.
      if xf_index != 0
        attributes << 's' << xf_index
      elsif @set_rows[row] && @set_rows[row][1]
        row_xf = @set_rows[row][1]
        attributes << 's' << row_xf.get_xf_index
      elsif @col_formats[col]
        col_xf = @col_formats[col]
        attributes << 's' << col_xf.get_xf_index
      end

      # Write the various cell types.
      case type
      when 'n'
        # Write a number.
        @writer.start_tag('c', attributes)
        write_cell_value(token)
        @writer.end_tag('c')
      when 's'
        # Write a string.
        attributes << 't' << 's'
        @writer.start_tag('c', attributes)
        write_cell_value(token)
        @writer.end_tag('c')
      when 'f'
        # Write a formula.
        @writer.start_tag('c', attributes)
        write_cell_formula(token)
        write_cell_value(cell[3] || 0)
        @writer.end_tag('c')
      when 'a'
        # Write an array formula.
        @writer.start_tag('c', attributes)
        write_cell_array_formula(token, cell[3])
        write_cell_value(cell[4])
        @writer.end_tag('c')
      when 'l'
        link_type = cell[3]

        # Write the string part a hyperlink.
        attributes << 't' << 's'
        @writer.start_tag('c', attributes)
        write_cell_value(token)
        @writer.end_tag('c')

        if link_type == 1
          # External link with rel file relationship.
          @hlink_count += 1
          @hlink_refs <<
            [
              link_type,    row,     col,
              @hlink_count, cell[5], cell[6]
            ]

          @external_hyper_links << [ '/hyperlink', cell[4], 'External' ]
        elsif link_type
          # External link with rel file relationship.
          @hlink_refs << [link_type, row, col, cell[4], cell[5], cell[6] ]
        end
      when 'b'
        # Write a empty cell.
        @writer.empty_tag('c', attributes)
      end
    end

    #
    # Write the cell value <v> element.
    #
    def write_cell_value(value = '')
      value ||= ''
      value = value.to_i if value == value.to_i
      @writer.data_element('v', value)
    end

    #
    # Write the cell formula <f> element.
    #
    def write_cell_formula(formula = '')
      @writer.data_element('f', formula)
    end

    #
    # Write the cell array formula <f> element.
    #
    def write_cell_array_formula(formula, range)
      attributes = ['t', 'array', 'ref', range]

      @writer.data_element('f', formula, attributes)
    end

    #
    # Write the frozen or split <pane> elements.
    #
    def write_panes
      return if @panes.empty?

      if @panes[4] == 2
        write_split_panes(*(@panes))
      else
        write_freeze_panes(*(@panes))
      end
    end

    #
    # Write the <pane> element for freeze panes.
    #
    def write_freeze_panes(row, col, top_row, left_col, type)
      y_split       = row
      x_split       = col
      top_left_cell = xl_rowcol_to_cell(top_row, left_col)

      # Move user cell selection to the panes.
      unless @selections.empty?
        dummy, active_cell, sqref = @selections[0]
        @selections = []
      end

      active_cell ||= nil
      sqref       ||= nil
      # Set the active pane.
      if row > 0 && col > 0
        active_pane = 'bottomRight'
        row_cell = xl_rowcol_to_cell(row, 0)
        col_cell = xl_rowcol_to_cell(0, col)
        @selections <<
            [ 'topRight',    col_cell,    col_cell ] <<
            [ 'bottomLeft',  row_cell,    row_cell ] <<
            [ 'bottomRight', active_cell, sqref ]
      elsif col > 0
        active_pane = 'topRight'
        @selections << [ 'topRight', active_cell, sqref ]
      else
        active_pane = 'bottomLeft'
        @selections << [ 'bottomLeft', active_cell, sqref ]
      end

      # Set the pane type.
      if type == 0
        state = 'frozen'
      elsif type == 1
        state = 'frozenSplit'
      else
        state = 'split'
      end

      attributes = []
      (attributes << 'xSplit' << x_split) if x_split > 0
      (attributes << 'ySplit' << y_split) if y_split > 0
      attributes << 'topLeftCell' << top_left_cell
      attributes << 'activePane'  << active_pane
      attributes << 'state'       << state

      @writer.empty_tag('pane', attributes)
    end

    #
    # Write the <pane> element for split panes.
    #
    # See also, implementers note for split_panes().
    #
    def write_split_panes(row, col, top_row, left_col, type)
      has_selection = false
      y_split = row
      x_split = col

      # Move user cell selection to the panes.
      if !@selections.empty?
        dummy, active_cell, sqref = @selections[0]
        @selections = []
        has_selection = true
      end

      # Convert the row and col to 1/20 twip units with padding.
      y_split = (20 * y_split + 300).to_i if y_split > 0
      x_split = calculate_x_split_width(x_split) if x_split > 0

      # For non-explicit topLeft definitions, estimate the cell offset based
      # on the pixels dimensions. This is only a workaround and doesn't take
      # adjusted cell dimensions into account.
      if top_row == row && left_col == col
        top_row  = (0.5 + ( y_split - 300 ) / 20 / 15).to_i
        left_col = (0.5 + ( x_split - 390 ) / 20 / 3 * 4 / 64).to_i
      end

      top_left_cell = xl_rowcol_to_cell(top_row, left_col)

      # If there is no selection set the active cell to the top left cell.
      if !has_selection
        active_cell = top_left_cell
        sqref       = top_left_cell
      end

      # Set the Cell selections.
      if row > 0 && col > 0
        active_pane = 'bottomRight'
        row_cell = xl_rowcol_to_cell(top_row, 0)
        col_cell = xl_rowcol_to_cell(0, left_col)

        @selections <<
          [ 'topRight',    col_cell,    col_cell ] <<
          [ 'bottomLeft',  row_cell,    row_cell ] <<
          [ 'bottomRight', active_cell, sqref ]
      elsif col > 0
        active_pane = 'topRight'
        @selections << [ 'topRight', active_cell, sqref ]
      else
        active_pane = 'bottomLeft'
        @selections << [ 'bottomLeft', active_cell, sqref ]
      end

      attributes = []
      (attributes << 'xSplit' << x_split) if x_split > 0
      (attributes << 'ySplit' << y_split) if y_split > 0
      attributes << 'topLeftCell' << top_left_cell
      (attributes << 'activePane' << active_pane ) if has_selection

      @writer.empty_tag('pane', attributes)
    end

    #
    # Convert column width from user units to pane split width.
    #
    def calculate_x_split_width(width)
      max_digit_width = 7    # For Calabri 11.
      padding         = 5

      # Convert to pixels.
      if width < 1
        pixels = int( width * 12 + 0.5 )
      else
        pixels = (width * max_digit_width + 0.5).to_i + padding
      end

      # Convert to points.
      points = pixels * 3 / 4

      # Convert to twips (twentieths of a point).
      twips = points * 20

      # Add offset/padding.
      twips + 390
    end

    #
    # Write the <sheetCalcPr> element for the worksheet calculation properties.
    #
    def write_sheet_calc_pr
      full_calc_on_load = 1

      attributes = ['fullCalcOnLoad', full_calc_on_load]

      @writer.empty_tag('sheetCalcPr', attributes)
    end

    #
    # Write the <phoneticPr> element.
    #
    def write_phonetic_pr
      font_id = 1
      type    = 'noConversion'

      attributes = [
          'fontId', font_id,
          'type',   type
      ]

      @writer.empty_tag('phoneticPr', attributes)
    end

    #
    # Write the <pageMargins> element.
    #
    def write_page_margins
      attributes = [
        'left',   @margin_left,
        'right',  @margin_right,
        'top',    @margin_top,
        'bottom', @margin_bottom,
        'header', @margin_header,
        'footer', @margin_footer
      ]
      @writer.empty_tag('pageMargins', attributes)
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
    def write_page_setup
      attributes = []

      return unless page_setup_changed?

      # Set paper size.
      attributes << 'paperSize' << @paper_size if @paper_size

      # Set the print_scale
      attributes << 'scale' << @print_scale if @print_scale != 100

      # Set the "Fit to page" properties.
      attributes << 'fitToWidth' << @fit_width if @fit_page && @fit_width != 1

      attributes << 'fitToHeight' << @fit_height if @fit_page && @fit_height != 1

      # Set the page print direction.
      attributes << 'pageOrder' << "overThenDown" if @page_order

      # Set page orientation.
      if orientation?
        attributes << 'orientation' << 'portrait'
      else
        attributes << 'orientation' << 'landscape'
      end

      @writer.empty_tag('pageSetup', attributes)
    end

    #
    # Write the <extLst> element.
    #
    def write_ext_lst
      @writer.start_tag('extLst')
      write_ext
      @writer.end_tag('extLst')
    end

    #
    # Write the <ext> element.
    #
    def write_ext
      xmlnsmx = 'http://schemas.microsoft.com/office/mac/excel/2008/main'
      uri     = 'http://schemas.microsoft.com/office/mac/excel/2008/main'

      attributes = [
        'xmlns:mx', xmlnsmx,
        'uri',      uri
      ]

      @writer.start_tag('ext', attributes)
      write_mx_plv
      @writer.end_tag('ext')
    end

    #
    # Write the <mx:PLV> element.
    #
    def write_mx_plv
      mode     = 1
      one_page = 0
      w_scale  = 0

      attributes = [
        'Mode',    mode,
        'OnePage', one_page,
        'WScale',  w_scale
      ]

      @writer.empty_tag('mx:PLV', attributes)
    end

    #
    # Write the <mergeCells> element.
    #
    def write_merge_cells
      return if @merge.empty?

      attributes = ['count', @merge.size]

      @writer.start_tag('mergeCells', attributes)

      # Write the mergeCell element.
      @merge.each { |merged_range| write_merge_cell(merged_range) }

      @writer.end_tag('mergeCells')
    end


    #
    # Write the <mergeCell> element.
    #
    def write_merge_cell(merged_range)
      row_min, col_min, row_max, col_max = merged_range

      # Convert the merge dimensions to a cell range.
      cell_1 = xl_rowcol_to_cell(row_min, col_min)
      cell_2 = xl_rowcol_to_cell(row_max, col_max)
      ref    = "#{cell_1}:#{cell_2}"

      attributes = ['ref', ref]

      @writer.empty_tag('mergeCell', attributes)
    end

    #
    # Write the <printOptions> element.
    #
    def write_print_options
      attributes = []

      return unless print_options_changed?

      # Set horizontal centering.
      attributes << 'horizontalCentered' << 1 if hcenter?

      # Set vertical centering.
      attributes << 'verticalCentered' << 1   if vcenter?

      # Enable row and column headers.
      attributes << 'headings' << 1 if print_headers?

      # Set printed gridlines.
      attributes << 'gridLines' << 1 if print_gridlines?

      @writer.empty_tag('printOptions', attributes)
    end

    #
    # Write the <headerFooter> element.
    #
    def write_header_footer
      return unless header_footer_changed?

      @writer.start_tag('headerFooter')
      write_odd_header if @header && @header != ''
      write_odd_footer if @footer && @footer != ''
      @writer.end_tag('headerFooter')
    end

    #
    # Write the <oddHeader> element.
    #
    def write_odd_header
      @writer.data_element('oddHeader', @header)
    end

    # _write_odd_footer()
    #
    # Write the <oddFooter> element.
    #
    def write_odd_footer
      @writer.data_element('oddFooter', @footer)
    end

    #
    # Write the <rowBreaks> element.
    #
    def write_row_breaks
      page_breaks = sort_pagebreaks(*(@hbreaks))
      count       = page_breaks.size

      return if page_breaks.empty?

      attributes = ['count', count, 'manualBreakCount', count]

      @writer.start_tag('rowBreaks', attributes)

      page_breaks.each { |row_num| write_brk(row_num, 16383) }

      @writer.end_tag('rowBreaks')
    end

    #
    # Write the <colBreaks> element.
    #
    def write_col_breaks
      page_breaks = sort_pagebreaks(*(@vbreaks))
      count       = page_breaks.size

      return if page_breaks.empty?

      attributes = ['count', count, 'manualBreakCount', count]

      @writer.start_tag('colBreaks', attributes)

      page_breaks.each { |col_num| write_brk(col_num, 1048575) }

      @writer.end_tag('colBreaks')
    end

    #
    # Write the <brk> element.
    #
    def write_brk(id, max)
      attributes = [
        'id',  id,
        'max', max,
        'man', 1
      ]

      @writer.empty_tag('brk', attributes)
    end

    #
    # Write the <autoFilter> element.
    #
    def write_auto_filter
      return unless autofilter_ref?

      attributes = ['ref', @autofilter_ref]

      if filter_on?
        # Autofilter defined active filters.
        @writer.start_tag('autoFilter', attributes)
        write_autofilters
        @writer.end_tag('autoFilter')
      else
        # Autofilter defined without active filters.
        @writer.empty_tag('autoFilter', attributes)
      end
    end

    #
    # Function to iterate through the columns that form part of an autofilter
    # range and write the appropriate filters.
    #
    def write_autofilters
      col1, col2 = @filter_range

      (col1 .. col2).each do |col|
        # Skip if column doesn't have an active filter.
        next unless @filter_cols[col]

        # Retrieve the filter tokens and write the autofilter records.
        tokens = @filter_cols[col]
        type   = @filter_type[col]

        write_filter_column(col, type, *tokens)
      end
    end

    #
    # Write the <filterColumn> element.
    #
    def write_filter_column(col_id, type, *filters)
      attributes = ['colId', col_id]

      @writer.start_tag('filterColumn', attributes)
      if type == 1
        # Type == 1 is the new XLSX style filter.
        write_filters(*filters)
      else
        # Type == 0 is the classic "custom" filter.
        write_custom_filters(*filters)
      end

      @writer.end_tag('filterColumn')
    end

    #
    # Write the <filters> element.
    #
    def write_filters(*filters)
      if filters.size == 1 && filters[0] == 'blanks'
        # Special case for blank cells only.
        @writer.empty_tag('filters', ['blank', 1])
      else
        # General case.
        @writer.start_tag('filters')
        filters.each { |filter| write_filter(filter) }
        @writer.end_tag('filters')
      end
    end

    #
    # Write the <filter> element.
    #
    def write_filter(val)
      @writer.empty_tag('filter', ['val', val])
    end


    #
    # Write the <customFilters> element.
    #
    def write_custom_filters(*tokens)
      if tokens.size == 2
        # One filter expression only.
        @writer.start_tag('customFilters')
        write_custom_filter(*tokens)
        @writer.end_tag('customFilters')
      else
        # Two filter expressions.

        # Check if the "join" operand is "and" or "or".
        if tokens[2] == 0
          attributes = ['and', 1]
        else
          attributes = ['and', 0]
        end

        # Write the two custom filters.
        @writer.start_tag('customFilters', attributes)
        write_custom_filter(tokens[0], tokens[1])
        write_custom_filter(tokens[3], tokens[4])
        @writer.end_tag('customFilters')
      end
    end


    #
    # Write the <customFilter> element.
    #
    def write_custom_filter(operator, val)
      operators = {
        1  => 'lessThan',
        2  => 'equal',
        3  => 'lessThanOrEqual',
        4  => 'greaterThan',
        5  => 'notEqual',
        6  => 'greaterThanOrEqual',
        22 => 'equal'
      }

      # Convert the operator from a number to a descriptive string.
      if operators[operator]
        operator = operators[operator]
      else
        raise "Unknown operator = #{operator}\n"
      end

      # The 'equal' operator is the default attribute and isn't stored.
      attributes = []
      attributes << 'operator' << operator unless operator == 'equal'
      attributes << 'val' << val

      @writer.empty_tag('customFilter', attributes)
    end

    #
    # Write the <hyperlinks> element. The attributes are different for internal
    # and external links.
    #
    def write_hyperlinks
      return if @hlink_refs.empty?

      @writer.start_tag('hyperlinks')

      @hlink_refs.each do |aref|
        type, *args = aref

        if type == 1
          write_hyperlink_external(*args)
        elsif type == 2
          write_hyperlink_internal(*args)
        end
      end

      @writer.end_tag('hyperlinks')
    end

    #
    # Write the <hyperlink> element for external links.
    #
    def write_hyperlink_external(row, col, id, location = nil, tooltip = nil)
      ref = xl_rowcol_to_cell(row, col)
      r_id = "rId#{id}"

      attributes = ['ref', ref, 'r:id', r_id]

      attributes << 'location' << location  if location
      attributes << 'tooltip'  << tooltip   if tooltip

      @writer.empty_tag('hyperlink', attributes)
    end

    #
    # Write the <hyperlink> element for internal links.
    #
    def write_hyperlink_internal(row, col, location, display, tooltip = nil)
      ref = xl_rowcol_to_cell(row, col)

      attributes = ['ref', ref, 'location', location]

      attributes << 'tooltip' << tooltip if tooltip
      attributes << 'display' << display

      @writer.empty_tag('hyperlink', attributes)
    end

    #
    # Write the <tabColor> element.
    #
    def write_tab_color
      return unless tab_color?

      attributes = ['rgb', get_palette_color(@tab_color)]
      @writer.empty_tag('tabColor', attributes)
    end

    #
    # Write the <sheetProtection> element.
    #
    def write_sheet_protection
      return unless protect?

      attributes = []
      attributes << "password"         << @protect[:password] if @protect[:password]
      attributes << "sheet"            << 1 if @protect[:sheet]
      attributes << "content"          << 1 if @protect[:content]
      attributes << "objects"          << 1 if !@protect[:objects]
      attributes << "scenarios"        << 1 if !@protect[:scenarios]
      attributes << "formatCells"      << 0 if @protect[:format_cells]
      attributes << "formatColumns"    << 0 if @protect[:format_columns]
      attributes << "formatRows"       << 0 if @protect[:format_rows]
      attributes << "insertColumns"    << 0 if @protect[:insert_columns]
      attributes << "insertRows"       << 0 if @protect[:insert_rows]
      attributes << "insertHyperlinks" << 0 if @protect[:insert_hyperlinks]
      attributes << "deleteColumns"    << 0 if @protect[:delete_columns]
      attributes << "deleteRows"       << 0 if @protect[:delete_rows]

      attributes << "selectLockedCells" << 1 if !@protect[:select_locked_cells]

      attributes << "sort"        << 0 if @protect[:sort]
      attributes << "autoFilter"  << 0 if @protect[:autofilter]
      attributes << "pivotTables" << 0 if @protect[:pivot_tables]

      attributes << "selectUnlockedCells" << 1 if !@protect[:select_unlocked_cells]

      @writer.empty_tag('sheetProtection', attributes)
    end

    #
    # Write the <drawing> elements.
    #
    def write_drawings
      write_drawing(@hlink_count + 1) if drawing?
    end

    #
    # Write the <drawing> element.
    #
    def write_drawing(id)
      r_id = "rId#{id}"

      attributes = ['r:id', r_id]

      @writer.empty_tag('drawing', attributes)
    end

    #
    # Write the <legacyDrawing> element.
    #
    def write_legacy_drawing
      return unless @has_comments

      # Increment the relationship id for any drawings or comments.
      id = @hlink_count + 1
      id += 1 if @drawing

      attributes = ['r:id', "rId#{id}"]

      @writer.empty_tag('legacyDrawing', attributes)
    end

    #
    # Write the <font> element.
    #
    def write_font(format)
      @rstring.start_tag('rPr')

      @rstring.empty_tag('b')       if !format.bold.nil? && format.bold != 0
      @rstring.empty_tag('i')       if !format.italic.nil? &&format.italic != 0
      @rstring.empty_tag('strike')  if !format.font_strikeout.nil? && format.font_strikeout != 0
      @rstring.empty_tag('outline') if !format.font_outline.nil? && format.font_outline != 0
      @rstring.empty_tag('shadow')  if !format.font_shadow && format.font_shadow != 0

      # Handle the underline variants.
      write_underline(format.underline) if !format.underline.nil? && format.underline != 0

      write_vert_align('superscript') if format.font_script == 1
      write_vert_align('subscript')   if format.font_script == 2

      @rstring.empty_tag('sz', ['val', format.size])

      theme = format.theme
      color = format.color
      if !theme.nil? && theme != 0
        write_color('theme', theme)
      elsif !color.nil? && color != 0
        color = get_palette_color(color)

        write_color('rgb', color)
      else
          write_color('theme', 1)
      end

      @rstring.empty_tag('rFont',  ['val', format.font])
      @rstring.empty_tag('family', ['val', format.font_family])

      if format.font == 'Calibri' && format.hyperlink == 0
        @rstring.empty_tag('scheme', ['val', format.font_scheme])
      end

      @rstring.end_tag('rPr')
    end

    #
    # Write the underline font element.
    #
    def write_underline(underline)
      # Handle the underline variants.
      if underline == 2
        attributes = [val, 'double']
      elsif underline == 33
        attributes = [val, 'singleAccounting']
      elsif underline == 34
        attributes = [val, 'doubleAccounting']
      else
        attributes = []    # Default to single underline.
      end

      @rstring.empty_tag('u', attributes)
    end

    #
    # Write the <vertAlign> font sub-element.
    #
    def write_vert_align(val)
      attributes = ['val', val]

      @rstring.empty_tag('vertAlign', attributes)
    end

    #
    # Write the <color> element.
    #
    def write_color(name, value)
      attributes = [name, value]

      @rstring.empty_tag('color', attributes)
    end

    #
    # Write the <dataValidations> element.
    #
    def write_data_validations
      return if @validations.empty?

      attributes = ['count', @validations.size]

      @writer.start_tag('dataValidations', attributes)
      @validations.each { |validation| write_data_validation(validation) }
      @writer.end_tag('dataValidations')
    end

    #
    # Write the <dataValidation> element.
    #
    def write_data_validation(param)
      sqref      = ''
      attributes = []

      # Set the cell range(s) for the data validation.
      param[:cells].each do |cells|
        # Add a space between multiple cell ranges.
        sqref += ' ' if sqref != ''

        row_first, col_first, row_last, col_last = cells

        # Swap last row/col for first row/col as necessary
        row_first, row_last = row_last, row_first if row_first > row_last
        col_first, col_last = col_last, col_first if col_first > col_last

        # If the first and last cell are the same write a single cell.
        if row_first == row_last && col_first == col_last
          sqref += xl_rowcol_to_cell(row_first, col_first)
        else
          sqref += xl_range(row_first, row_last, col_first, col_last)
        end
      end

      #use Data::Dumper::Perltidy
      #print Dumper param

      attributes << 'type' << param[:validate]
      attributes << 'operator' << param[:criteria] if param[:criteria] != 'between'

      if param[:error_type]
        attributes << 'errorStyle' << 'warning' if param[:error_type] == 1
        attributes << 'errorStyle' << 'information' if param[:error_type] == 2
      end
      attributes << 'allowBlank'       << 1 if param[:ignore_blank] != 0
      attributes << 'showDropDown'     << 1 if param[:dropdown]     == 0
      attributes << 'showInputMessage' << 1 if param[:show_input]   != 0
      attributes << 'showErrorMessage' << 1 if param[:show_error]   != 0

      attributes << 'errorTitle' << param[:error_title]  if param[:error_title]
      attributes << 'error' << param[:error_message]     if param[:error_message]
      attributes << 'promptTitle' << param[:input_title] if param[:input_title]
      attributes << 'prompt' << param[:input_message]    if param[:input_message]
      attributes << 'sqref' << sqref

      @writer.start_tag('dataValidation', attributes)

      # Write the formula1 element.
      write_formula_1(param[:value])

      # Write the formula2 element.
      write_formula_2(param[:maximum]) if param[:maximum]

      @writer.end_tag('dataValidation')
    end

    #
    # Write the <formula1> element.
    #
    def write_formula_1(formula)
      # Convert a list array ref into a comma separated string.
      formula   = %!"#{formula.join(',')}"! if formula.kind_of?(Array)

      formula = formula.sub(/^=/, '') if formula.respond_to?(:sub)

      @writer.data_element('formula1', formula)
    end

    # _write_formula_2()
    #
    # Write the <formula2> element.
    #
    def write_formula_2(formula)
      formula = formula.sub(/^=/, '') if formula.respond_to?(:sub)

      @writer.data_element('formula2', formula)
    end

    # in Perl module : _write_formula()
    #
    def write_formula_tag(data)
      @writer.data_element('formula', data)
    end

    #
    # Write the Worksheet conditional formats.
    #
    def write_conditional_formats
      ranges = @cond_formats.keys.sort
      return if ranges.empty?

      ranges.each { |range| write_conditional_formatting(range, @cond_formats[range]) }
    end

    #
    # Write the <conditionalFormatting> element.
    #
    def write_conditional_formatting(range, params)
      attributes = ['sqref', range]

      @writer.start_tag('conditionalFormatting', attributes)

      params.each { |param| write_cf_rule(param) }

      @writer.end_tag('conditionalFormatting')
    end

    #
    # Write the <cfRule> element.
    #
    def write_cf_rule(param)
      attributes = ['type' , param[:type]]

      if param[:format]
        attributes << 'dxfId' << param[:format]
      end
      attributes << 'priority' << param[:priority]
      attributes << 'operator' << param[:criteria]

      @writer.start_tag('cfRule', attributes)

      if param[:type] == 'cellIs'
        if param[:minimum] && param[:maximum]
          write_formula_tag(param[:minimum])
          write_formula_tag(param[:maximum])
        else
          write_formula_tag(param[:value])
        end
      end

      @writer.end_tag('cfRule')
    end

    def store_data_to_table(row, col, data)
      if @table[row]
        @table[row][col] = data
      else
        @table[row] = []
        @table[row][col] = data
      end
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
    # Check that $row and $col are valid and store max and min values for use in
    # DIMENSIONS record. See, store_dimensions().
    #
    # The $ignore_row/$ignore_col flags is used to indicate that we wish to
    # perform the dimension check without storing the value.
    #
    # The ignore flags are use by set_row() and data_validate.
    #
    def check_dimensions(row, col, ignore_row = 0, ignore_col = 0)       #:nodoc:
      return -2 unless row
      return -2 if row >= @xls_rowmax

      return -2 unless col
      return -2 if col >= @xls_colmax

      if ignore_row == 0
        @dim_rowmin = row if !@dim_rowmin || (row < @dim_rowmin)
        @dim_rowmax = row if !@dim_rowmax || (row > @dim_rowmax)
      end

      if ignore_col == 0
        @dim_colmin = col if !@dim_colmin || (col < @dim_colmin)
        @dim_colmax = col if !@dim_colmax || (col > @dim_colmax)
      end

      0
    end

    #
    # Calculate the "spans" attribute of the <row> tag. This is an XLSX
    # optimisation and isn't strictly required. However, it makes comparing
    # files easier.
    #
    # The span is the same for each block of 16 rows.
    #
    def calculate_spans
      span_min = nil
      span_max = 0
      spans = []
      (@dim_rowmin .. @dim_rowmax).each do |row_num|
        row_ref = @table[row_num]
        if row_ref
          (@dim_colmin .. @dim_colmax).each do |col_num|
            col_ref = @table[row_num][col_num]
            if col_ref
              if !span_min
                span_min = col_num
                span_max = col_num
              else
                span_min = col_num if col_num < span_min
                span_max = col_num if col_num > span_max
              end
            end
          end
        end

        if ((row_num + 1) % 16 == 0) || (row_num == @dim_rowmax)
          span_index = row_num / 16
          if span_min
            span_min += 1
            span_max += 1
            spans[span_index] = "#{span_min}:#{span_max}"
            span_min = nil
          end
        end
      end

      @row_spans = spans
    end

    def xf(format)
      if format.kind_of?(Format)
        format.xf_index
      else
        0
      end
    end

    #
    # Add a string to the shared string table, if it isn't already there, and
    # return the string index.
    #
    def shared_string_index(str)
      # Add the string to the shared string table.
      unless @workbook.str_table[str]
        @workbook.str_table[str] = @workbook.str_unique
        @workbook.str_unique += 1
      end

      @workbook.str_total += 1
      @workbook.str_table[str]
    end

    #
    # convert_name_area(first_row, first_col, last_row, last_col)
    #
    # Convert zero indexed rows and columns to the format required by worksheet
    # named ranges, eg, "Sheet1!$A$1:$C$13".
    #
    def convert_name_area(row_num_1, col_num_1, row_num_2, col_num_2)
      range1       = ''
      range2       = ''
      row_col_only = false

      # Convert to A1 notation.
      col_char_1 = xl_col_to_name(col_num_1, 1)
      col_char_2 = xl_col_to_name(col_num_2, 1)
      row_char_1 = "$#{row_num_1 + 1}"
      row_char_2 = "$#{row_num_2 + 1}"

      # We need to handle some special cases that refer to rows or columns only.
      if row_num_1 == 0 and row_num_2 == @xls_rowmax - 1
        range1       = col_char_1
        range2       = col_char_2
        row_col_only = true
      elsif col_num_1 == 0 and col_num_2 == @xls_colmax - 1
        range1       = row_char_1
        range2       = row_char_2
        row_col_only = true
      else
        range1 = col_char_1 + row_char_1
        range2 = col_char_2 + row_char_2
      end

      # A repeated range is only written once (if it isn't a special case).
      if range1 == range2 && !row_col_only
        area = range1
      else
        area = "#{range1}:#{range2}"
      end

      # Build up the print area range "Sheet1!$A$1:$C$13".
      "#{quote_sheetname(name)}!#{area}"
    end

    #
    # Sheetnames used in references should be quoted if they contain any spaces,
    # special characters or if the look like something that isn't a sheet name.
    # TODO. We need to handle more special cases.
    #
    def quote_sheetname(sheetname)
      return sheetname if sheetname =~ /^Sheet\d+$/
      return "'#{sheetname}'"
    end

    def fit_page?
      if @fit_page
        @fit_page != 0
      else
        false
      end
    end

    def filter_on?
      if @filter_on
        @filter_on != 0
      else
        false
      end
    end

    def tab_color?
      if @tab_color
        @tab_color != 0
      else
        false
      end
    end

    def zoom_scale_normal?
      !!@zoom_scale_normal
    end

    def page_view?
      !!@page_view
    end

    def right_to_left?
      !!@right_to_left
    end

    def show_zeros?
      !!@show_zeros
    end

    def screen_gridlines?
      !!@screen_gridlines
    end

    def protect?
      !!@protect
    end

    def autofilter_ref?
      !!@autofilter_ref
    end

    def date_1904?
      @workbook.date_1904?
    end

    def print_options_changed?
      !!@print_options_changed
    end

    def hcenter?
      !!@hcenter
    end

    def vcenter?
      !!@vcenter
    end

    def print_headers?
      !!@print_headers
    end

    def print_gridlines?
      !!@print_gridlines
    end

    def page_setup_changed?
      !!@page_setup_changed
    end

    def orientation?
      !!@orientation
    end

    def header_footer_changed?
      !!@header_footer_changed
    end

    def drawing?
      !!@drawing
    end

    def remove_white_space(margin)
      if margin.respond_to?(:gsub)
        margin.gsub(/[^\d\.]/, '')
      else
        margin
      end
    end
  end
end
