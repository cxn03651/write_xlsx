# -*- coding: utf-8 -*-
require 'write_xlsx/package/xml_writer_simple'
require 'write_xlsx/colors'
require 'write_xlsx/format'
require 'write_xlsx/drawing'
require 'write_xlsx/compatibility'
require 'write_xlsx/utility'
require 'tempfile'

module Writexlsx
  #
  # A new worksheet is created by calling the add_worksheet() method from a workbook object:
  #
  #     worksheet1 = workbook.add_worksheet
  #     worksheet2 = workbook.add_worksheet
  #
  # The following methods are available through a new worksheet:
  #
  #     write
  #     write_number
  #     write_string
  #     write_rich_string
  #     write_blank
  #     write_row
  #     write_col
  #     write_date_time
  #     write_url
  #     write_url_range
  #     write_formula
  #     write_comment
  #     show_comments
  #     set_comments_author
  #     insert_image
  #     insert_chart
  #     data_validation
  #     conditional_format
  #     get_name
  #     activate
  #     select
  #     hide
  #     set_first_sheet
  #     protect
  #     set_selection
  #     set_row
  #     set_column
  #     outline_settings
  #     freeze_panes
  #     split_panes
  #     merge_range
  #     merge_range_type
  #     set_zoom
  #     right_to_left
  #     hide_zero
  #     set_tab_color
  #     autofilter
  #     filter_column
  #     filter_column_list
  #
  # ==Cell notation
  #
  # WriteXLSX supports two forms of notation to designate the position of cells:
  # Row-column notation and A1 notation.
  #
  # Row-column notation uses a zero based index for both row and column
  # while A1 notation uses the standard Excel alphanumeric sequence of column letter
  # and 1-based row. For example:
  #
  #     (0, 0)      # The top left cell in row-column notation.
  #     ('A1')      # The top left cell in A1 notation.
  #
  #     (1999, 29)  # Row-column notation.
  #     ('AD2000')  # The same cell in A1 notation.
  #
  # Row-column notation is useful if you are referring to cells programmatically:
  #
  #     (0..9).each do |i|
  #       worksheet.write(i, 0, 'Hello')    # Cells A1 to A10
  #     end
  #
  # A1 notation is useful for setting up a worksheet manually and
  # for working with formulas:
  #
  #     worksheet.write('H1', 200)
  #     worksheet.write('H2', '=H1+1')
  #
  # In formulas and applicable methods you can also use the A:A column notation:
  #
  #     worksheet.write('A1', '=SUM(B:B)')
  #
  # The Writexlsx::Utility module that is included in the distro contains
  # helper functions for dealing with A1 notation, for example:
  #
  #     include Writexlsx::Utility
  #
  #     row, col = xl_cell_to_rowcol('C2')    # (1, 2)
  #     str      = xl_rowcol_to_cell(1, 2)    # C2
  #
  # For simplicity, the parameter lists for the worksheet method calls in the
  # following sections are given in terms of row-column notation. In all cases
  # it is also possible to use A1 notation.
  #
  class Worksheet
    include Writexlsx::Utility

    RowMax   = 1048576  # :nodoc:
    ColMax   = 16384    # :nodoc:
    StrMax   = 32767    # :nodoc:
    Buffer   = 4096     # :nodoc:

    attr_writer :fit_page
    attr_reader :index, :_repeat_cols, :_repeat_rows
    attr_reader :charts, :images, :drawing
    attr_reader :external_hyper_links, :external_drawing_links, :external_comment_links, :drawing_links
    attr_reader :vml_data_id, :vml_shape_id, :comments_array
    attr_reader :autofilter_area, :hidden

    def initialize(workbook, index, name) #:nodoc:
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

    def set_xml_writer(filename) #:nodoc:
      @writer.set_xml_writer(filename)
    end

    def assemble_xml_file #:nodoc:
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

    #
    # The name() method is used to retrieve the name of a worksheet.
    # For example:
    #
    #     workbook.sheets.each do |sheet|
    #       print sheet.name
    #     end
    #
    # For reasons related to the design of WriteXLSX and to the internals
    # of Excel there is no set_name() method. The only way to set the
    # worksheet name is via the add_worksheet() method.
    #
    def name
      @name
    end

    #
    # Set this worksheet as a selected worksheet, i.e. the worksheet has its tab
    # highlighted.
    #
    # The select() method is used to indicate that a worksheet is selected in
    # a multi-sheet workbook:
    #
    #     worksheet1.activate
    #     worksheet2.select
    #     worksheet3.select
    #
    # A selected worksheet has its tab highlighted. Selecting worksheets is a
    # way of grouping them together so that, for example, several worksheets
    # could be printed in one go. A worksheet that has been activated via
    # the activate() method will also appear as selected.
    #
    def select
      @hidden   = false  # Selected worksheet can't be hidden.
      @selected = true
    end

    #
    # Set this worksheet as the active worksheet, i.e. the worksheet that is
    # displayed when the workbook is opened. Also set it as selected.
    #
    # The activate() method is used to specify which worksheet is initially
    # visible in a multi-sheet workbook:
    #
    #     worksheet1 = workbook.add_worksheet('To')
    #     worksheet2 = workbook.add_worksheet('the')
    #     worksheet3 = workbook.add_worksheet('wind')
    #
    #     worksheet3.activate
    #
    # This is similar to the Excel VBA activate method. More than one worksheet
    # can be selected via the select() method, see below, however only one
    # worksheet can be active.
    #
    # The default active worksheet is the first worksheet.
    #
    def activate
      @hidden = false
      @selected = true
      @workbook.activesheet = @index
    end

    #
    # Hide this worksheet.
    #
    # The hide() method is used to hide a worksheet:
    #
    #     worksheet2.hide
    #
    # You may wish to hide a worksheet in order to avoid confusing a user
    # with intermediate data or calculations.
    #
    # A hidden worksheet can not be activated or selected so this method
    # is mutually exclusive with the activate() and select() methods. In
    # addition, since the first worksheet will default to being the active
    # worksheet, you cannot hide the first worksheet without activating another
    # sheet:
    #
    #     worksheet2.activate
    #     worksheet1.hide
    #
    def hide
      @hidden = true
      @selected = false
      @workbook.activesheet = 0
      @workbook.firstsheet  = 0
    end

    def hidden? # :nodoc:
      @hidden
    end

    #
    # Set this worksheet as the first visible sheet. This is necessary
    # when there are a large number of worksheets and the activated
    # worksheet is not visible on the screen.
    #
    # The activate() method determines which worksheet is initially selected.
    # However, if there are a large number of worksheets the selected
    # worksheet may not appear on the screen. To avoid this you can select
    # which is the leftmost visible worksheet using set_first_sheet():
    #
    #     20.times { workbook.add_worksheet }
    #
    #     worksheet21 = workbook.add_worksheet
    #     worksheet22 = workbook.add_worksheet
    #
    #     worksheet21.set_first_sheet
    #     worksheet22.activate
    #
    # This method is not required very often. The default value is the first worksheet.
    #
    def set_first_sheet
      @hidden = false
      @workbook.firstsheet = self
    end

    #
    # Set the worksheet protection flags to prevent modification of worksheet
    # objects.
    #
    # The protect() method is used to protect a worksheet from modification:
    #
    #     worksheet.protect
    #
    # The protect() method also has the effect of enabling a cell's locked
    # and hidden properties if they have been set. A locked cell cannot be
    # edited and this property is on by default for all cells. A hidden
    # cell will display the results of a formula but not the formula itself.
    #
    # See the protection.rb program in the examples directory of the distro
    # for an illustrative example and the set_locked and set_hidden format
    # methods in "CELL FORMATTING".
    #
    # You can optionally add a password to the worksheet protection:
    #
    #     worksheet.protect('drowssap')
    #
    # Passing the empty string '' is the same as turning on protection
    # without a password.
    #
    # Note, the worksheet level password in Excel provides very weak
    # protection. It does not encrypt your data and is very easy to
    # deactivate. Full workbook encryption is not supported by WriteXLSX
    # since it requires a completely different file format and would take
    # several man months to implement.
    #
    # You can specify which worksheet elements that you which to protect
    # by passing a hash_ref with any or all of the following keys:
    #
    #     # Default shown.
    #     options = {
    #         :objects               => false,
    #         :scenarios             => false,
    #         :format_cells          => false,
    #         :format_columns        => false,
    #         :format_rows           => false,
    #         :insert_columns        => false,
    #         :insert_rows           => false,
    #         :insert_hyperlinks     => false,
    #         :delete_columns        => false,
    #         :delete_rows           => false,
    #         :select_locked_cells   => true,
    #         :sort                  => false,
    #         :autofilter            => false,
    #         :pivot_tables          => false,
    #         :select_unlocked_cells => true
    #     }
    # The default boolean values are shown above. Individual elements
    # can be protected as follows:
    #
    #     worksheet.protect('drowssap', { :insert_rows => true } )
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
    # :call-seq:
    #   set_column(firstcol, lastcol, width, format, hidden, level)
    #
    # This method can be used to change the default properties of a single
    # column or a range of columns. All parameters apart from first_col
    # and last_col are optional.
    #
    # If set_column() is applied to a single column the value of first_col
    # and last_col should be the same. In the case where $last_col is zero
    # it is set to the same value as first_col.
    #
    # It is also possible, and generally clearer, to specify a column range
    # using the form of A1 notation used for columns. See the note about
    # "Cell notation".
    #
    # Examples:
    #
    #     worksheet.set_column(0, 0, 20)    # Column  A   width set to 20
    #     worksheet.set_column(1, 3, 30)    # Columns B-D width set to 30
    #     worksheet.set_column('E:E', 20)   # Column  E   width set to 20
    #     worksheet.set_column('F:H', 30)   # Columns F-H width set to 30
    #
    # The width corresponds to the column width value that is specified in
    # Excel. It is approximately equal to the length of a string in the
    # default font of Arial 10. Unfortunately, there is no way to specify
    # "AutoFit" for a column in the Excel file format. This feature is
    # only available at runtime from within Excel.
    #
    # As usual the format parameter is optional, for additional information,
    # see "CELL FORMATTING". If you wish to set the format without changing
    # the width you can pass nil as the width parameter:
    #
    #     worksheet.set_column(0, 0, nil, format)
    #
    # The format parameter will be applied to any cells in the column that
    # don't have a format. For example
    #
    #     worksheet.set_column( 'A:A', nil, format1 )    # Set format for col 1
    #     worksheet.write( 'A1', 'Hello' )                  # Defaults to format1
    #     worksheet.write( 'A2', 'Hello', format2 )        # Keeps format2
    #
    # If you wish to define a column format in this way you should call the
    # method before any calls to write(). If you call it afterwards it
    # won't have any effect.
    #
    # A default row format takes precedence over a default column format
    #
    #     worksheet.set_row( 0, nil, format1 )           # Set format for row 1
    #     worksheet.set_column( 'A:A', nil, format2 )    # Set format for col 1
    #     worksheet.write( 'A1', 'Hello' )               # Defaults to format1
    #     worksheet.write( 'A2', 'Hello' )               # Defaults to format2
    #
    # The hidden parameter should be set to 1 if you wish to hide a column.
    # This can be used, for example, to hide intermediary steps in a
    # complicated calculation:
    #
    #     worksheet.set_column( 'D:D', 20,  format, 1 )
    #     worksheet.set_column( 'E:E', nil, nil,    1 )
    #
    # The level parameter is used to set the outline level of the column.
    # Outlines are described in "OUTLINES AND GROUPING IN EXCEL". Adjacent
    # columns with the same outline level are grouped together into a single
    # outline.
    #
    # The following example sets an outline level of 1 for columns B to G:
    #
    #     worksheet.set_column( 'B:G', nil, nil, 0, 1 )
    #
    # The hidden parameter can also be used to hide collapsed outlined
    # columns when used in conjunction with the level parameter.
    #
    #     worksheet.set_column( 'B:G', nil, nil, 1, 1 )
    #
    # For collapsed outlines you should also indicate which row has the
    # collapsed + symbol using the optional collapsed parameter.
    #
    #     worksheet.set_column( 'H:H', nil, nil, 0, 0, 1 )
    #
    # For a more complete example see the outline.rb and outline_collapsed.rb
    # programs in the examples directory of the distro.
    #
    # Excel allows up to 7 outline levels. Therefore the level parameter
    # should be in the range 0 <= level <= 7.
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
      width  ||= 0                        # Ensure width isn't nil.
      width = 0 if hidden && hidden != 0  # Set width to zero if col is hidden

      (firstcol .. lastcol).each do |col|
        @col_sizes[col]   = width
        @col_formats[col] = format if format
      end
    end

    #
    # :call-seq:
    #   set_selection(cell_or_cell_range)
    #
    # Set which cell or cells are selected in a worksheet.
    #
    # This method can be used to specify which cell or cells are selected
    # in a worksheet. The most common requirement is to select a single cell,
    # in which case $last_row and $last_col can be omitted. The active cell
    # within a selected range is determined by the order in which first and
    # last are specified. It is also possible to specify a cell or a range
    # using A1 notation. See the note about "Cell notation".
    #
    # Examples:
    #
    #     worksheet1.set_selection(3, 3)          # 1. Cell D4.
    #     worksheet2.set_selection(3, 3, 6, 6)    # 2. Cells D4 to G7.
    #     worksheet3.set_selection(6, 6, 3, 3)    # 3. Cells G7 to D4.
    #     worksheet4.set_selection('D4')          # Same as 1.
    #     worksheet5.set_selection('D4:G7')       # Same as 2.
    #     worksheet6.set_selection('G7:D4')       # Same as 3.
    #
    # The default cell selections is (0, 0), 'A1'.
    #
    def set_selection(*args)
      return if args.empty?

      args = row_col_notation(args)

      # There should be either 2 or 4 arguments.
      case args.size
      when 2
        # Single cell selection.
        active_cell = xl_rowcol_to_cell(args[0], args[1])
        sqref = active_cell
      when 4
        # Range selection.
        active_cell = xl_rowcol_to_cell(args[0], args[1])

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
    # :call-seq:
    #   freeze_panes(row, col [ , top_row, left_col ] )
    #
    # This method can be used to divide a worksheet into horizontal or
    # vertical regions known as panes and to also "freeze" these panes so
    # that the splitter bars are not visible. This is the same as the
    # Window->Freeze Panes menu command in Excel
    #
    # The parameters $row and $col are used to specify the location of
    # the split. It should be noted that the split is specified at the
    # top or left of a cell and that the method uses zero based indexing.
    # Therefore to freeze the first row of a worksheet it is necessary
    # to specify the split at row 2 (which is 1 as the zero-based index).
    # This might lead you to think that you are using a 1 based index
    # but this is not the case.
    #
    # You can set one of the $row and $col parameters as zero if you
    # do not want either a vertical or horizontal split.
    #
    # Examples:
    #
    #     worksheet.freeze_panes(1, 0)    # Freeze the first row
    #     worksheet.freeze_panes('A2')    # Same using A1 notation
    #     worksheet.freeze_panes(0, 1)    # Freeze the first column
    #     worksheet.freeze_panes('B1')    # Same using A1 notation
    #     worksheet.freeze_panes(1, 2)    # Freeze first row and first 2 columns
    #     worksheet.freeze_panes('C2')    # Same using A1 notation
    #
    # The parameters $top_row and $left_col are optional. They are used
    # to specify the top-most or left-most visible row or column in the
    # scrolling region of the panes. For example to freeze the first row
    # and to have the scrolling region begin at row twenty:
    #
    #     worksheet.freeze_panes(1, 0, 20, 0)
    #
    # You cannot use A1 notation for the $top_row and $left_col parameters.
    #
    # See also the panes.pl program in the examples directory of the
    # distribution.
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
    # :call-seq:
    #   split_panes(y, x, top_row, left_col, offset_row, offset_col)
    #
    # Set panes and mark them as split.
    #--
    # Implementers note. The API for this method doesn't map well from the XLS
    # file format and isn't sufficient to describe all cases of split panes.
    # It should probably be something like:
    #
    #     split_panes($y, $x, $top_row, $left_col, $offset_row, $offset_col)
    #
    # I'll look at changing this if it becomes an issue.
    #++
    # This method can be used to divide a worksheet into horizontal or vertical
    # regions known as panes. This method is different from the freeze_panes()
    # method in that the splits between the panes will be visible to the user
    # and each pane will have its own scroll bars.
    #
    # The parameters $y and $x are used to specify the vertical and horizontal
    # position of the split. The units for $y and $x are the same as those
    # used by Excel to specify row height and column width. However, the
    # vertical and horizontal units are different from each other. Therefore
    # you must specify the $y and $x parameters in terms of the row heights
    # and column widths that you have set or the default values which are 15
    # for a row and 8.43 for a column.
    #
    # You can set one of the $y and $x parameters as zero if you do not want
    # either a vertical or horizontal split. The parameters top_row and left_col
    # are optional. They are used to specify the top-most or left-most visible
    # row or column in the bottom-right pane.
    #
    # Example:
    #
    #     worksheet.split_panes(15, 0   )    # First row
    #     worksheet.split_panes( 0, 8.43)    # First column
    #     worksheet.split_panes(15, 8.43)    # First row and column
    #
    # You cannot use A1 notation with this method.
    #
    # See also the freeze_panes() method and the panes.rb program in the
    # examples directory of the distribution.
    #
    def split_panes(*args)
      # Call freeze panes but add the type flag for split panes.
      freeze_panes(args[0], args[1], args[2], args[3], 2)
    end

    #
    # Set the page orientation as portrait.
    # The default worksheet orientation is portrait, so you won't generally
    # need to call this method.
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
    # This method is used to display the worksheet in "Page View/Layout" mode.
    #
    def set_page_view(flag = true)
      @page_view = !!flag
    end

    #
    # Set the colour of the worksheet tab.
    #
    # The set_tab_color() method is used to change the colour of the worksheet
    # tab. This feature is only available in Excel 2002 and later. You can use
    # one of the standard colour names provided by the Format object or a
    # colour index. See "COLOURS IN EXCEL" and the set_custom_color() method.
    #
    #     worksheet1.set_tab_color('red')
    #     worksheet2.set_tab_color(0x0C)
    #
    # See the tab_colors.pl program in the examples directory of the distro.
    #
    def set_tab_color(color)
      @tab_color = Colors.new.get_color(color)
    end

    #
    # Set the paper type. Ex. 1 = US Letter, 9 = A4
    #
    # This method is used to set the paper format for the printed output of
    # a worksheet. The following paper styles are available:
    #
    #     Index   Paper format            Paper size
    #     =====   ============            ==========
    #       0     Printer default         -
    #       1     Letter                  8 1/2 x 11 in
    #       2     Letter Small            8 1/2 x 11 in
    #       3     Tabloid                 11 x 17 in
    #       4     Ledger                  17 x 11 in
    #       5     Legal                   8 1/2 x 14 in
    #       6     Statement               5 1/2 x 8 1/2 in
    #       7     Executive               7 1/4 x 10 1/2 in
    #       8     A3                      297 x 420 mm
    #       9     A4                      210 x 297 mm
    #      10     A4 Small                210 x 297 mm
    #      11     A5                      148 x 210 mm
    #      12     B4                      250 x 354 mm
    #      13     B5                      182 x 257 mm
    #      14     Folio                   8 1/2 x 13 in
    #      15     Quarto                  215 x 275 mm
    #      16     -                       10x14 in
    #      17     -                       11x17 in
    #      18     Note                    8 1/2 x 11 in
    #      19     Envelope  9             3 7/8 x 8 7/8
    #      20     Envelope 10             4 1/8 x 9 1/2
    #      21     Envelope 11             4 1/2 x 10 3/8
    #      22     Envelope 12             4 3/4 x 11
    #      23     Envelope 14             5 x 11 1/2
    #      24     C size sheet            -
    #      25     D size sheet            -
    #      26     E size sheet            -
    #      27     Envelope DL             110 x 220 mm
    #      28     Envelope C3             324 x 458 mm
    #      29     Envelope C4             229 x 324 mm
    #      30     Envelope C5             162 x 229 mm
    #      31     Envelope C6             114 x 162 mm
    #      32     Envelope C65            114 x 229 mm
    #      33     Envelope B4             250 x 353 mm
    #      34     Envelope B5             176 x 250 mm
    #      35     Envelope B6             176 x 125 mm
    #      36     Envelope                110 x 230 mm
    #      37     Monarch                 3.875 x 7.5 in
    #      38     Envelope                3 5/8 x 6 1/2 in
    #      39     Fanfold                 14 7/8 x 11 in
    #      40     German Std Fanfold      8 1/2 x 12 in
    #      41     German Legal Fanfold    8 1/2 x 13 in
    #
    # Note, it is likely that not all of these paper types will be available
    # to the end user since it will depend on the paper formats that the
    # user's printer supports. Therefore, it is best to stick to standard
    # paper types.
    #
    #     worksheet.set_paper(1)    # US Letter
    #     worksheet.set_paper(9)    # A4
    #
    # If you do not specify a paper type the worksheet will print using
    # the printer's default paper.
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
    # Headers and footers are generated using a string which is a combination
    # of plain text and control characters. The margin parameter is optional.
    #
    # The available control character are:
    #
    #     Control             Category            Description
    #     =======             ========            ===========
    #     &L                  Justification       Left
    #     &C                                      Center
    #     &R                                      Right
    #
    #     &P                  Information         Page number
    #     &N                                      Total number of pages
    #     &D                                      Date
    #     &T                                      Time
    #     &F                                      File name
    #     &A                                      Worksheet name
    #     &Z                                      Workbook path
    #
    #     &fontsize           Font                Font size
    #     &"font,style"                           Font name and style
    #     &U                                      Single underline
    #     &E                                      Double underline
    #     &S                                      Strikethrough
    #     &X                                      Superscript
    #     &Y                                      Subscript
    #
    #     &&                  Miscellaneous       Literal ampersand &
    #
    # Text in headers and footers can be justified (aligned) to the left,
    # center and right by prefixing the text with the control characters
    # &L, &C and &R.
    #
    # For example (with ASCII art representation of the results):
    #
    #     worksheet.set_header('&LHello')
    #
    #      ---------------------------------------------------------------
    #     |                                                               |
    #     | Hello                                                         |
    #     |                                                               |
    #
    #
    #     worksheet.set_header('&CHello')
    #
    #      ---------------------------------------------------------------
    #     |                                                               |
    #     |                          Hello                                |
    #     |                                                               |
    #
    #
    #     worksheet.set_header('&RHello')
    #
    #      ---------------------------------------------------------------
    #     |                                                               |
    #     |                                                         Hello |
    #     |                                                               |
    #
    # For simple text, if you do not specify any justification the text will
    # be centred. However, you must prefix the text with &C if you specify
    # a font name or any other formatting:
    #
    #     worksheet.set_header('Hello')
    #
    #      ---------------------------------------------------------------
    #     |                                                               |
    #     |                          Hello                                |
    #     |                                                               |
    #
    # You can have text in each of the justification regions:
    #
    #     worksheet.set_header('&LCiao&CBello&RCielo')
    #
    #      ---------------------------------------------------------------
    #     |                                                               |
    #     | Ciao                     Bello                          Cielo |
    #     |                                                               |
    #
    # The information control characters act as variables that Excel will update
    # as the workbook or worksheet changes. Times and dates are in the users
    # default format:
    #
    #     worksheet.set_header('&CPage &P of &N')
    #
    #      ---------------------------------------------------------------
    #     |                                                               |
    #     |                        Page 1 of 6                            |
    #     |                                                               |
    #
    #
    #     worksheet.set_header('&CUpdated at &T')
    #
    #      ---------------------------------------------------------------
    #     |                                                               |
    #     |                    Updated at 12:30 PM                        |
    #     |                                                               |
    #
    # You can specify the font size of a section of the text by prefixing it
    # with the control character &n where n is the font size:
    #
    #     worksheet1.set_header('&C&30Hello Big' )
    #     worksheet2.set_header('&C&10Hello Small' )
    #
    # You can specify the font of a section of the text by prefixing it with
    # the control sequence &"font,style" where fontname is a font name such
    # as "Courier New" or "Times New Roman" and style is one of the standard
    # Windows font descriptions: "Regular", "Italic", "Bold" or "Bold Italic":
    #
    #     worksheet1.set_header('&C&"Courier New,Italic"Hello')
    #     worksheet2.set_header('&C&"Courier New,Bold Italic"Hello')
    #     worksheet3.set_header('&C&"Times New Roman,Regular"Hello')
    #
    # It is possible to combine all of these features together to create
    # sophisticated headers and footers. As an aid to setting up complicated
    # headers and footers you can record a page set-up as a macro in Excel
    # and look at the format strings that VBA produces. Remember however
    # that VBA uses two double quotes "" to indicate a single double quote.
    # For the last example above the equivalent VBA code looks like this:
    #
    #     .LeftHeader   = ""
    #     .CenterHeader = "&""Times New Roman,Regular""Hello"
    #     .RightHeader  = ""
    #
    # To include a single literal ampersand & in a header or footer you
    # should use a double ampersand &&:
    #
    #     worksheet1.set_header('&CCuriouser && Curiouser - Attorneys at Law')
    #
    # As stated above the margin parameter is optional. As with the other
    # margins the value should be in inches. The default header and footer
    # margin is 0.3 inch. Note, the default margin is different from the
    # default used in the binary file format by Spreadsheet::WriteExcel.
    # The header and footer margin size can be set as follows:
    #
    #     worksheet.set_header('&CHello', 0.75)
    #
    # The header and footer margins are independent of the top and bottom
    # margins.
    #
    # Note, the header or footer string must be less than 255 characters.
    # Strings longer than this will not be written and a warning will be
    # generated.
    #
    # See, also the headers.rb program in the examples directory of the
    # distribution.
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
    # The syntax of the set_footer() method is the same as set_header()
    #
    def set_footer(string = '', margin = 0.3)
      raise 'Footer string must be less than 255 characters' if string.length >= 255

      @footer                = string
      @margin_footer         = margin
      @header_footer_changed = true
    end

    #
    # Center the worksheet data horizontally between the margins on the printed page:
    #
    def center_horizontally
      @print_options_changed = true
      @hcenter               = true
    end

    #
    # Center the worksheet data vertically between the margins on the printed page:
    #
    def center_vertically
      @print_options_changed = true
      @vcenter               = true
    end

    #
    # Set all the page margins to the same value in inches.
    #
    # There are several methods available for setting the worksheet margins
    # on the printed page:
    #
    #     set_margins()        # Set all margins to the same value
    #     set_margins_LR()     # Set left and right margins to the same value
    #     set_margins_TB()     # Set top and bottom margins to the same value
    #     set_margin_left()    # Set left margin
    #     set_margin_right()   # Set right margin
    #     set_margin_top()     # Set top margin
    #     set_margin_bottom()  # Set bottom margin
    #
    # All of these methods take a distance in inches as a parameter.
    # Note: 1 inch = 25.4mm. ;-) The default left and right margin is 0.7 inch.
    # The default top and bottom margin is 0.75 inch. Note, these defaults
    # are different from the defaults used in the binary file format
    # by writeexcel gem.
    #
    def set_margins(margin)
      set_margin_left(margin)
      set_margin_right(margin)
      set_margin_top(margin)
      set_margin_bottom(margin)
    end

    #
    # Set the left and right margins to the same value in inches.
    # See set_margins
    #
    def set_margins_LR(margin)
      set_margin_left(margin)
      set_margin_right(margin)
    end

    #
    # Set the top and bottom margins to the same value in inches.
    # See set_margins
    #
    def set_margins_TB(margin)
      set_margin_top(margin)
      set_margin_bottom(margin)
    end

    #
    # Set the left margin in inches.
    # See set_margins
    #
    def set_margin_left(margin = 0.7)
      @margin_left = remove_white_space(margin)
    end

    #
    # Set the right margin in inches.
    # See set_margins
    #
    def set_margin_right(margin = 0.7)
      @margin_right = remove_white_space(margin)
    end

    #
    # Set the top margin in inches.
    # See set_margins
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
    # Set the number of rows to repeat at the top of each printed page.
    #
    # For large Excel documents it is often desirable to have the first row
    # or rows of the worksheet print out at the top of each page. This can
    # be achieved by using the repeat_rows() method. The parameters
    # first_row and last_row are zero based. The last_row parameter is
    # optional if you only wish to specify one row:
    #
    #     worksheet1.repeat_rows(0)    # Repeat the first row
    #     worksheet2.repeat_rows(0, 1) # Repeat the first two rows
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
    # :call-seq:
    #   repeat_columns(first_col, last_col = nil)
    #
    # Set the columns to repeat at the left hand side of each printed page.
    #
    # For large Excel documents it is often desirable to have the first
    # column or columns of the worksheet print out at the left hand side
    # of each page. This can be achieved by using the repeat_columns()
    # method. The parameters first_column and last_column are zero based.
    # The last_column parameter is optional if you only wish to specify
    # one column. You can also specify the columns using A1 column
    # notation, see the note about "Cell notation".
    #
    #     worksheet1.repeat_columns(0)        # Repeat the first column
    #     worksheet2.repeat_columns(0, 1)     # Repeat the first two columns
    #     worksheet3.repeat_columns('A:A')    # Repeat the first column
    #     worksheet4.repeat_columns('A:B')    # Repeat the first two columns
    #
    def repeat_columns(*args)
      if args[0] =~ /^\D/
        args = substitute_cellref(args)
        # Returned values $row1 and $row2 aren't required here. Remove them.
        args = [args[1], args[3]]
      end

      col_min = args[0]
      col_max = args[1] || args[0]

      # Convert to A notation.
      col_min = xl_col_to_name(args[0], 1)
      col_max = xl_col_to_name(args[1], 1)

      area = col_min +  ':' + col_max

      # Build up the print area range "=Sheet2!C1:C2"
      sheetname = quote_sheetname(@name)
      area = sheetname + "!" + area

      @repeat_cols = area
    end

    def get_print_area # :nodoc:
      @print_area.dup
    end

    #
    # :call-seq:
    #   print_area(first_row, first_col, last_row, last_col)
    #
    # This method is used to specify the area of the worksheet that will
    # be printed. All four parameters must be specified. You can also use
    # A1 notation, see the note about "Cell notation".
    #
    #     $worksheet1->print_area( 'A1:H20' );    # Cells A1 to H20
    #     $worksheet2->print_area( 0, 0, 19, 7 ); # The same
    #     $worksheet2->print_area( 'A:H' );       # Columns A to H if rows have data
    #
    def print_area(*args)
      return @print_area if args.empty?
      
      args = substitute_cellref(args) if args[0] =~ /^\D/
      # Check for a cell reference in A1 notation and substitute row and column
      return if args.size != 4    # Require 4 parameters

      row1, col1, row2, col2 = args

      # Ignore max print area since this is the same as no print area for Excel.
      if row1 == 0 && col1 == 0 && row2 == @xls_rowmax - 1 && col2 == @xls_colmax - 1
        return
      end

      # Build up the print area range "=Sheet2!R1C1:R2C1"
      @print_area = convert_name_area(row1, col1, row2, col2)
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
    # Set the scale factor of the printed page.
    # Scale factors in the range 10 <= scale <= 400 are valid:
    #
    #     worksheet1.set_print_scale( 50)
    #     worksheet2.set_print_scale( 75)
    #     worksheet3.set_print_scale(300)
    #     worksheet4.set_print_scale(400)
    #
    # The default scale factor is 100. Note, set_print_scale() does not
    # affect the scale of the visible page in Excel. For that you should
    # use set_zoom().
    #
    # Note also that although it is valid to use both fit_to_pages() and
    # set_print_scale() on the same worksheet only one of these options
    # can be active at a time. The last method call made will set
    # the active option.
    #
    def set_print_scale(scale = 100)
      # Confine the scale to Excel's range
      scale = 100 if scale < 10 || scale > 400

      # Turn off "fit to page" option.
      @fit_page = 0

      @print_scale        = scale.to_i
      @page_setup_changed = 1
    end

    #
    # Display the worksheet right to left for some eastern versions of Excel.
    #
    # The right_to_left() method is used to change the default direction
    # of the worksheet from left-to-right, with the A1 cell in the top
    # left, to right-to-left, with the he A1 cell in the top right.
    #
    #     worksheet.right_to_left
    #
    # This is useful when creating Arabic, Hebrew or other near or far
    # eastern worksheets that use right-to-left as the default direction.
    #
    def right_to_left(flag = true)
      @right_to_left = !!flag
    end

    #
    # Hide cell zero values.
    #
    # The hide_zero() method is used to hide any zero values that appear
    # in cells.
    #
    #     worksheet.hide_zero
    #
    # In Excel this option is found under Tools->Options->View.
    #
    def hide_zero(flag = true)
        @show_zeros = !flag
    end

    #
    # Set the order in which pages are printed.
    #
    # The print_across method is used to change the default print direction.
    # This is referred to by Excel as the sheet "page order".
    #
    #     worksheet.print_across
    #
    # The default page order is shown below for a worksheet that extends
    # over 4 pages. The order is called "down then across":
    #
    #     [1] [3]
    #     [2] [4]
    #
    # However, by using the print_across method the print order will be
    # changed to "across then down":
    #
    #     [1] [2]
    #     [3] [4]
    #
    def print_across(page_order = true)
      if page_order
        @page_order         = true
        @page_setup_changed = true
      else
        @page_order = false
      end
    end

    #
    # The set_start_page() method is used to set the number of the
    # starting page when the worksheet is printed out.
    # The default value is 1.
    #
    #     worksheet.set_start_page(2)
    #
    def set_start_page(page_start)
      @page_start   = page_start
      @custom_start = 1
    end

    #
    # :call-seq:
    #  write(row, column [ , token [ , format ] ])
    #
    # Excel makes a distinction between data types such as strings, numbers,
    # blanks, formulas and hyperlinks. To simplify the process of writing
    # data the write() method acts as a general alias for several more
    # specific methods:
    #
    #     write_string
    #     write_number
    #     write_blank
    #     write_formula
    #     write_url
    #     write_row
    #     write_col
    #
    # The general rule is that if the data looks like a something then
    # a something is written. Here are some examples in both row-column
    # and A1 notation:
    #
    #                                                     # Same as:
    #     worksheet.write(0, 0, 'Hello'                ) # write_string()
    #     worksheet.write(1, 0, 'One'                  ) # write_string()
    #     worksheet.write(2, 0,  2                     ) # write_number()
    #     worksheet.write(3, 0,  3.00001               ) # write_number()
    #     worksheet.write(4, 0,  ""                    ) # write_blank()
    #     worksheet.write(5, 0,  ''                    ) # write_blank()
    #     worksheet.write(6, 0,  nil                   ) # write_blank()
    #     worksheet.write(7, 0                         ) # write_blank()
    #     worksheet.write(8, 0,  'http://www.ruby.com/') # write_url()
    #     worksheet.write('A9',  'ftp://ftp.ruby.org/' ) # write_url()
    #     worksheet.write('A10', 'internal:Sheet1!A1'  ) # write_url()
    #     worksheet.write('A11', 'external:c:\foo.xlsx') # write_url()
    #     worksheet.write('A12', '=A3 + 3*A4'          ) # write_formula()
    #     worksheet.write('A13', '=SIN(PI()/4)'        ) # write_formula()
    #     worksheet.write('A14', [1, 2]                ) # write_row()
    #     worksheet.write('A15', [ [1, 2] ]            ) # write_col()
    #
    #     # Write an array formula. Not available in writeexcel gem.
    #     worksheet.write('A16', '{=SUM(A1:B1*A2:B2)}' ) # write_formula()
    #
    # The format parameter is optional. It should be a valid Format object.
    #
    #     format = workbook.add_format
    #     format.set_bold
    #     format.set_color('red')
    #     format.set_align('center')
    #
    #     worksheet.write(4, 0, 'Hello', $format)    # Formatted string
    #
    # The write() method will ignore empty strings or nil tokens unless a format
    # is also supplied. As such you needn't worry about special handling for
    # empty or nil in your data. See also the write_blank() method.
    #
    # One problem with the write() method is that occasionally data looks like
    # a number but you don't want it treated as a number. For example, zip
    # codes or ID numbers often start with a leading zero.
    # If you want to write this data with leading zero(s), use write_string.
    #
    # The write methods return:
    #     0 for success.
    #    -1 for insufficient number of arguments.
    #    -2 for row or column out of bounds.
    #    -3 for string too long.
    #
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
    # :call-seq:
    #   write_row(row, col, array [ , format ] )
    #
    # Write a row of data starting from (row, col). Call write_col() if any of
    # the elements of the array are in turn array. This allows the writing
    # of 1D or 2D arrays of data in one go.
    #
    # Returns: the first encountered error value or zero for no errors
    #
    # The write_row() method can be used to write a 1D or 2D array of data
    # in one go. This is useful for converting the results of a database
    # query into an Excel worksheet. You must pass a reference to the array
    # of data rather than the array itself. The write() method is then
    # called for each element of the data. For example:
    #
    #     array = ['awk', 'gawk', 'mawk']
    #
    #     worksheet.write_row(0, 0, array)
    #
    #     # The above example is equivalent to:
    #     worksheet.write(0, 0, array[0])
    #     worksheet.write(0, 1, array[1])
    #     worksheet.write(0, 2, array[2])
    #
    # Note: For convenience the write() method behaves in the same way as
    # write_row() if it is passed an array reference.
    # Therefore the following two method calls are equivalent:
    #
    #     worksheet.write_row('A1', array)    # Write a row of data
    #     worksheet.write(    'A1', array)    # Same thing
    #
    # As with all of the write methods the format parameter is optional.
    # If a format is specified it is applied to all the elements of the
    # data array.
    #
    # Array references within the data will be treated as columns.
    # This allows you to write 2D arrays of data in one go. For example:
    #
    #     eec =  [
    #                 ['maggie', 'milly', 'molly', 'may'  ],
    #                 [13,       14,      15,      16     ],
    #                 ['shell',  'star',  'crab',  'stone']
    #            ]
    #
    #     worksheet.write_row('A1', eec)
    # Would produce a worksheet as follows:
    #
    #      -----------------------------------------------------------
    #     |   |    A    |    B    |    C    |    D    |    E    | ...
    #      -----------------------------------------------------------
    #     | 1 | maggie  | 13      | shell   | ...     |  ...    | ...
    #     | 2 | milly   | 14      | star    | ...     |  ...    | ...
    #     | 3 | molly   | 15      | crab    | ...     |  ...    | ...
    #     | 4 | may     | 16      | stone   | ...     |  ...    | ...
    #     | 5 | ...     | ...     | ...     | ...     |  ...    | ...
    #     | 6 | ...     | ...     | ...     | ...     |  ...    | ...
    #
    # To write the data in a row-column order refer to the write_col()
    # method below.
    #
    # Any nil in the data will be ignored unless a format is applied to
    # the data, in which case a formatted blank cell will be written.
    # In either case the appropriate row or column value will still
    # be incremented.
    #
    # The write_row() method returns the first error encountered when
    # writing the elements of the data or zero if no errors were
    # encountered. See the return values described for the write()
    # method.
    #
    # See also the write_arrays.rb program in the examples directory
    # of the distro.
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
    # :call-seq:
    #   write_col(row, col, array [ , format ] )
    #
    # Write a column of data starting from (row, col). Call write_row() if any of
    # the elements of the array are in turn array. This allows the writing
    # of 1D or 2D arrays of data in one go.
    #
    # Returns: the first encountered error value or zero for no errors
    #
    # The write_col() method can be used to write a 1D or 2D array of data
    # in one go. This is useful for converting the results of a database
    # query into an Excel worksheet. You must pass a reference to the array
    # of data rather than the array itself. The write() method is then
    # called for each element of the data. For example:
    #
    #     array = [ 'awk', 'gawk', 'mawk' ]
    #
    #     worksheet.write_col(0, 0, array)
    #
    #     # The above example is equivalent to:
    #     worksheet.write(0, 0, array[0])
    #     worksheet.write(1, 0, array[1])
    #     worksheet.write(2, 0, array[2])
    #
    # As with all of the write methods the format parameter is optional.
    # If a format is specified it is applied to all the elements of the
    # data array.
    #
    # Array references within the data will be treated as rows.
    # This allows you to write 2D arrays of data in one go. For example:
    #
    #     eec =  [
    #                 ['maggie', 'milly', 'molly', 'may'  ],
    #                 [13,       14,      15,      16     ],
    #                 ['shell',  'star',  'crab',  'stone']
    #            ]
    #
    #     worksheet.write_col('A1', eec)
    #
    # Would produce a worksheet as follows:
    #
    #      -----------------------------------------------------------
    #     |   |    A    |    B    |    C    |    D    |    E    | ...
    #      -----------------------------------------------------------
    #     | 1 | maggie  | milly   | molly   | may     |  ...    | ...
    #     | 2 | 13      | 14      | 15      | 16      |  ...    | ...
    #     | 3 | shell   | star    | crab    | stone   |  ...    | ...
    #     | 4 | ...     | ...     | ...     | ...     |  ...    | ...
    #     | 5 | ...     | ...     | ...     | ...     |  ...    | ...
    #     | 6 | ...     | ...     | ...     | ...     |  ...    | ...
    #
    # To write the data in a column-row order refer to the write_row()
    # method above.
    #
    # Any nil in the data will be ignored unless a format is applied to
    # the data, in which case a formatted blank cell will be written.
    # In either case the appropriate row or column value will still be
    # incremented.
    #
    # As noted above the write() method can be used as a synonym for
    # write_row() and write_row() handles nested array refs as columns.
    # Therefore, the following two method calls are equivalent although
    # the more explicit call to write_col() would be preferable for
    # maintainability:
    #
    #     worksheet.write_col('A1', array     ) # Write a column of data
    #     worksheet.write(    'A1', [ array ] ) # Same thing
    #
    # The write_col() method returns the first error encountered when
    # writing the elements of the data or zero if no errors were encountered.
    # See the return values described for the write() method above.
    #
    # See also the write_arrays.rb program in the examples directory of
    # the distro.
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
    # :call-seq:
    #   write_comment(row, column, string, options = {})
    #
    # Write a comment to the specified row and column (zero indexed).
    #
    # write_comment methods return:
    #   Returns  0 : normal termination
    #           -1 : insufficient number of arguments
    #           -2 : row or column out of range
    #
    # The write_comment() method is used to add a comment to a cell.
    # A cell comment is indicated in Excel by a small red triangle in the
    # upper right-hand corner of the cell. Moving the cursor over the red
    # triangle will reveal the comment.
    #
    # The following example shows how to add a comment to a cell:
    #
    #     worksheet.write(        2, 2, 'Hello')
    #     worksheet.write_comment(2, 2, 'This is a comment.')
    #
    # As usual you can replace the row and column parameters with an A1
    # cell reference. See the note about "Cell notation".
    #
    #     worksheet.write(        'C3', 'Hello')
    #     worksheet.write_comment('C3', 'This is a comment.')
    #
    # The write_comment() method will also handle strings in UTF-8 format.
    #
    #     worksheet.write_comment('C3', "\x{263a}")       # Smiley
    #     worksheet.write_comment('C4', 'Comment ca va?')
    #
    # In addition to the basic 3 argument form of write_comment() you can
    # pass in several optional key/value pairs to control the format of
    # the comment. For example:
    #
    #     worksheet.write_comment('C3', 'Hello', :visible => 1, :author => 'Perl')
    #
    # Most of these options are quite specific and in general the default
    # comment behaviour will be all that you need. However, should you
    # need greater control over the format of the cell comment the
    # following options are available:
    #
    #     :author
    #     :visible
    #     :x_scale
    #     :width
    #     :y_scale
    #     :height
    #     :color
    #     :start_cell
    #     :start_row
    #     :start_col
    #     :x_offset
    #     :y_offset
    #
    # ===Option: author
    #
    # This option is used to indicate who is the author of the cell
    # comment. Excel displays the author of the comment in the status
    # bar at the bottom of the worksheet. This is usually of interest
    # in corporate environments where several people might review and
    # provide comments to a workbook.
    #
    #     worksheet.write_comment('C3', 'Atonement', :author => 'Ian McEwan')
    #
    # The default author for all cell comments can be set using the
    # set_comments_author() method.
    #
    #     worksheet.set_comments_author('Ruby')
    #
    # ===Option: visible
    #
    # This option is used to make a cell comment visible when the worksheet
    # is opened. The default behaviour in Excel is that comments are
    # initially hidden. However, it is also possible in Excel to make
    # individual or all comments visible. In WriteXLSX individual
    # comments can be made visible as follows:
    #
    #     worksheet.write_comment('C3', 'Hello', :visible => 1 )
    #
    # It is possible to make all comments in a worksheet visible
    # using the show_comments() worksheet method. Alternatively, if all of
    # the cell comments have been made visible you can hide individual comments:
    #
    #     worksheet.write_comment('C3', 'Hello', :visible => 0)
    #
    # ===Option: x_scale
    #
    # This option is used to set the width of the cell comment box as a
    # factor of the default width.
    #
    #     worksheet.write_comment('C3', 'Hello', :x_scale => 2)
    #     worksheet.write_comment('C4', 'Hello', :x_scale => 4.2)
    #
    # ===Option: width
    #
    # This option is used to set the width of the cell comment box
    # explicitly in pixels.
    #
    #     worksheet.write_comment('C3', 'Hello', :width => 200)
    #
    # ===Option: y_scale
    #
    # This option is used to set the height of the cell comment box as a
    # factor of the default height.
    #
    #     worksheet.write_comment('C3', 'Hello', :y_scale => 2)
    #     worksheet.write_comment('C4', 'Hello', :y_scale => 4.2)
    #
    # ===Option: height
    #
    # This option is used to set the height of the cell comment box
    # explicitly in pixels.
    #
    #     worksheet.write_comment('C3', 'Hello', :height => 200)
    #
    # ===Option: color
    #
    # This option is used to set the background colour of cell comment
    # box. You can use one of the named colours recognised by WriteXLSX
    # or a colour index. See "COLOURS IN EXCEL".
    #
    #     worksheet.write_comment('C3', 'Hello', :color => 'green')
    #     worksheet.write_comment('C4', 'Hello', :color => 0x35)      # Orange
    #
    # ===Option: start_cell
    #
    # This option is used to set the cell in which the comment will appear.
    # By default Excel displays comments one cell to the right and one cell
    # above the cell to which the comment relates. However, you can change
    # this behaviour if you wish. In the following example the comment
    # which would appear by default in cell D2 is moved to E2.
    #
    #     worksheet.write_comment('C3', 'Hello', :start_cell => 'E2')
    #
    # ===Option: start_row
    #
    # This option is used to set the row in which the comment will appear.
    # See the start_cell option above. The row is zero indexed.
    #
    #     worksheet.write_comment('C3', 'Hello', :start_row => 0)
    #
    # ===Option: start_col
    #
    # This option is used to set the column in which the comment will appear.
    # See the start_cell option above. The column is zero indexed.
    #
    #     worksheet.write_comment('C3', 'Hello', :start_col => 4)
    #
    # ===Option: x_offset
    #
    # This option is used to change the x offset, in pixels, of a comment
    # within a cell:
    #
    #     worksheet.write_comment('C3', $comment, :x_offset => 30)
    #
    # ===Option: y_offset
    #
    # This option is used to change the y offset, in pixels, of a comment
    # within a cell:
    #
    #     worksheet.write_comment('C3', $comment, :x_offset => 30)
    #
    # You can apply as many of these options as you require.
    #
    # Note about using options that adjust the position of the cell comment
    # such as start_cell, start_row, start_col, x_offset and y_offset:
    # Excel only displays offset cell comments when they are displayed as
    # "visible". Excel does not display hidden cells as moved when you
    # mouse over them.
    #
    # Note about row height and comments. If you specify the height of a
    # row that contains a comment then Excel::Writer::XLSX will adjust the
    # height of the comment to maintain the default or user specified
    # dimensions. However, the height of a row can also be adjusted
    # automatically by Excel if the text wrap property is set or large
    # fonts are used in the cell. This means that the height of the row
    # is unknown to the module at run time and thus the comment box is
    # stretched with the row. Use the set_row() method to specify the
    # row height explicitly and avoid this problem.
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

    #
    # :call-seq:
    #   write_number(row, column, number [ , format ] )
    #
    # Write an integer or a float to the cell specified by row and column:
    #
    #     worksheet.write_number(0, 0, 123456)
    #     worksheet.write_number('A2', 2.3451)
    #
    # See the note about "Cell notation".
    # The format parameter is optional.
    #
    # In general it is sufficient to use the write() method.
    #
    # Note: some versions of Excel 2007 do not display the calculated values
    # of formulas written by WriteXLSX. Applying all available Service Packs
    # to Excel should fix this.
    #
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
    # :call-seq:
    #   write_string(row, column, string [, format ] )
    #
    # Write a string to the specified row and column (zero indexed).
    # format is optional.
    #
    #   Returns  0 : normal termination
    #           -1 : insufficient number of arguments
    #           -2 : row or column out of range
    #           -3 : long string truncated to 32767 chars
    #
    # write_string methods return:
    #
    #     worksheet.write_string(0, 0, 'Your text here')
    #     worksheet.write_string('A2', 'or here')
    #
    # The maximum string size is 32767 characters. However the maximum
    # string segment that Excel can display in a cell is 1000.
    # All 32767 characters can be displayed in the formula bar.
    #
    # In general it is sufficient to use the write() method.
    # However, you may sometimes wish to use the write_string() method
    # to write data that looks like a number but that you don't want
    # treated as a number. For example, zip codes or phone numbers:
    #
    #     # Write as a plain string
    #     worksheet.write_string('A1', '01209')
    #
    # However, if the user edits this string Excel may convert it back
    # to a number. To get around this you can use the Excel text format @:
    #
    #     # Format as a string. Doesn't change to a number when edited
    #     format1 = workbook.add_format(:num_format => '@')
    #     worksheet.write_string('A2', '01209', format1)
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

    #
    # :call-seq:
    #    write_rich_string(row, column, format, string,  [,cell_format] )
    #
    # The write_rich_string() method is used to write strings with multiple formats.
    # The method receives string fragments prefixed by format objects. The final
    # format object is used as the cell format.
    #
    # write_rich_string methods return:
    #
    #   Returns  0 : normal termination.
    #           -1 : insufficient number of arguments.
    #           -2 : row or column out of range.
    #           -3 : long string truncated to 32767 chars.
    #           -4 : 2 consecutive formats used.
    #
    # For example to write the string "This is bold and this is italic"
    # you would use the following:
    #
    #     bold   = workbook.add_format(:bold   => 1)
    #     italic = workbook.add_format(:italic => 1)
    #
    #     worksheet.write_rich_string('A1',
    #         'This is ', bold, 'bold', ' and this is ', italic, 'italic')
    #
    # The basic rule is to break the string into fragments and put a format
    # object before the fragment that you want to format. For example:
    #
    #     # Unformatted string.
    #       'This is an example string'
    #
    #     # Break it into fragments.
    #       'This is an ', 'example', ' string'
    #
    #     # Add formatting before the fragments you want formatted.
    #       'This is an ', format, 'example', ' string'
    #
    #     # In WriteXLSX.
    #     worksheet.write_rich_string('A1',
    #         'This is an ', format, 'example', ' string')
    # String fragments that don't have a format are given a default
    # format. So for example when writing the string "Some bold text"
    # you would use the first example below but it would be equivalent
    # to the second:
    #
    #     # With default formatting:
    #     bold    = workbook.add_format(:bold => 1)
    #
    #     worksheet.write_rich_string('A1',
    #         'Some ', bold, 'bold', ' text')
    #
    #     # Or more explicitly:
    #     bold    = workbook.add_format(:bold => 1)
    #     default = workbook.add_format
    #
    #     worksheet.write_rich_string('A1',
    #         default, 'Some ', bold, 'bold', default, ' text')
    #
    # As with Excel, only the font properties of the format such as font
    # name, style, size, underline, color and effects are applied to the
    # string fragments. Other features such as border, background and
    # alignment must be applied to the cell.
    #
    # The write_rich_string() method allows you to do this by using the
    # last argument as a cell format (if it is a format object).
    # The following example centers a rich string in the cell:
    #
    #     bold   = workbook.add_format(:bold  => 1)
    #     center = workbook.add_format(:align => 'center')
    #
    #     worksheet.write_rich_string('A5',
    #         'Some ', bold, 'bold text', ' centered', center)
    #
    # See the rich_strings.rb example in the distro for more examples.
    #
    #     bold   = workbook.add_format(:bold        => 1)
    #     italic = workbook.add_format(:italic      => 1)
    #     red    = workbook.add_format(:color       => 'red')
    #     blue   = workbook.add_format(:color       => 'blue')
    #     center = workbook.add_format(:align       => 'center')
    #     super  = workbook.add_format(:font_script => 1)
    #
    #     # Write some strings with multiple formats.
    #     worksheet.write_rich_string('A1',
    #         'This is ', bold, 'bold', ' and this is ', italic, 'italic')
    #
    #     worksheet.write_rich_string('A3',
    #         'This is ', red, 'red', ' and this is ', blue, 'blue')
    #
    #     worksheet.write_rich_string('A5',
    #         'Some ', bold, 'bold text', ' centered', center)
    #
    #     worksheet.write_rich_string('A7',
    #         italic, 'j = k', super, '(n-1)', center)
    #
    # As with write_sting() the maximum string size is 32767 characters.
    # See also the note about "Cell notation".
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
    # :call-seq:
    #   write_blank(row, col, format)
    #
    # Write a blank cell to the specified row and column (zero indexed).
    # A blank cell is used to specify formatting without adding a string
    # or a number.
    #
    # A blank cell without a format serves no purpose. Therefore, we don't write
    # a BLANK record unless a format is specified. This is mainly an optimisation
    # for the write_row() and write_col() methods.
    #
    # write_blank methods return:
    #   Returns  0 : normal termination (including no format)
    #           -1 : insufficient number of arguments
    #           -2 : row or column out of range
    #
    # Excel differentiates between an "Empty" cell and a "Blank" cell.
    # An "Empty" cell is a cell which doesn't contain data whilst a "Blank"
    # cell is a cell which doesn't contain data but does contain formatting.
    # Excel stores "Blank" cells but ignores "Empty" cells.
    #
    # As such, if you write an empty cell without formatting it is ignored:
    #
    #     worksheet.write('A1', nil, format )    # write_blank()
    #     worksheet.write('A2', nil )            # Ignored
    #
    # This seemingly uninteresting fact means that you can write arrays of
    # data without special treatment for nil or empty string values.
    #
    # See the note about "Cell notation".
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

    #
    # :call-seq:
    #   write_formula(row, column, formula [ , format [ , value ] ] )
    #
    # Write a formula or function to the cell specified by row and column:
    #
    #     worksheet.write_formula(0, 0, '=$B$3 + B4')
    #     worksheet.write_formula(1, 0, '=SIN(PI()/4)')
    #     worksheet.write_formula(2, 0, '=SUM(B1:B5)')
    #     worksheet.write_formula('A4', '=IF(A3>1,"Yes", "No")')
    #     worksheet.write_formula('A5', '=AVERAGE(1, 2, 3, 4)')
    #     worksheet.write_formula('A6', '=DATEVALUE("1-Jan-2001")')
    # Array formulas are also supported:
    #
    #     worksheet.write_formula('A7', '{=SUM(A1:B1*A2:B2)}')
    #
    # See also the write_array_formula() method.
    #
    # See the note about "Cell notation". For more information about
    # writing Excel formulas see "FORMULAS AND FUNCTIONS IN EXCEL"
    #
    # If required, it is also possible to specify the calculated value
    # of the formula. This is occasionally necessary when working with
    # non-Excel applications that don't calculate the value of the
    # formula. The calculated value is added at the end of the argument list:
    #
    #     worksheet.write('A1', '=2+2', format, 4)
    #
    # However, this probably isn't something that will ever need to do.
    # If you do use this feature then do so with care.
    #
    def write_formula(*args)
      # Check for a cell reference in A1 notation and substitute row and column
      args = row_col_notation(args)

      return -1 if args.size < 3   # Check the number of args

      row, col, formula, format, value = args

      if formula =~ /^\{=.*\}$/
        return write_array_formula(row, col, row, col, formula, format, value)
      end

      # Check that row and col are valid and store max and min values
      return -2 unless check_dimensions(row, col) == 0

      formula.sub!(/^=/, '')

      store_data_to_table(row, col, ['f', formula, format, value])
      0
    end

    #
    # :call-seq:
    #   write_array_formula(row1, col1, row2, col2, formula [ , format [ , value ] ] )
    #
    # Write an array formula to the specified row and column (zero indexed).
    #
    # format is optional.
    #
    # write_array_formula methods return:
    #   Returns  0 : normal termination
    #           -1 : insufficient number of arguments
    #           -2 : row or column out of range
    #
    # In Excel an array formula is a formula that performs a calculation
    # on a set of values. It can return a single value or a range of values.
    #
    # An array formula is indicated by a pair of braces around the
    # formula: {=SUM(A1:B1*A2:B2)}. If the array formula returns a single
    # value then the first and last parameters should be the same:
    #
    #     worksheet.write_array_formula('A1:A1', '{=SUM(B1:C1*B2:C2)}')
    #
    # It this case however it is easier to just use the write_formula()
    # or write() methods:
    #
    #     # Same as above but more concise.
    #     worksheet.write('A1', '{=SUM(B1:C1*B2:C2)}')
    #     worksheet.write_formula('A1', '{=SUM(B1:C1*B2:C2)}')
    #
    # For array formulas that return a range of values you must specify
    # the range that the return values will be written to:
    #
    #     worksheet.write_array_formula('A1:A3',    '{=TREND(C1:C3,B1:B3)}')
    #     worksheet.write_array_formula(0, 0, 2, 0, '{=TREND(C1:C3,B1:B3)}')
    #
    # If required, it is also possible to specify the calculated value of
    # the formula. This is occasionally necessary when working with non-Excel
    # applications that don't calculate the value of the formula.
    # The calculated value is added at the end of the argument list:
    #
    #     worksheet.write_array_formula('A1:A3', '{=TREND(C1:C3,B1:B3)}', format, 105)
    #
    # In addition, some early versions of Excel 2007 don't calculate the
    # values of array formulas when they aren't supplied. Installing the
    # latest Office Service Pack should fix this issue.
    #
    # See also the array_formula.rb program in the examples directory of
    # the distro.
    #
    # Note: Array formulas are not supported by writeexcel gem.
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

    # The outline_settings() method is used to control the appearance of
    # outlines in Excel. Outlines are described in "OUTLINES AND GROUPING IN EXCEL".
    #
    # The visible parameter is used to control whether or not outlines are
    # visible. Setting this parameter to 0 will cause all outlines on the
    # worksheet to be hidden. They can be unhidden in Excel by means of the
    # "Show Outline Symbols" command button. The default setting is 1 for
    # visible outlines.
    #
    #     worksheet.outline_settings(0)
    #
    # The symbols_below parameter is used to control whether the row outline
    # symbol will appear above or below the outline level bar. The default
    # setting is 1 for symbols to appear below the outline level bar.
    #
    # The symbols_right parameter is used to control whether the column
    # outline symbol will appear to the left or the right of the outline level
    # bar. The default setting is 1 for symbols to appear to the right of
    # the outline level bar.
    #
    # The auto_style parameter is used to control whether the automatic
    # outline generator in Excel uses automatic styles when creating an
    # outline. This has no effect on a file generated by WriteXLSX but it
    # does have an effect on how the worksheet behaves after it is created.
    # The default setting is 0 for "Automatic Styles" to be turned off.
    #
    # The default settings for all of these parameters correspond to Excel's
    # default parameters.
    #
    # The worksheet parameters controlled by outline_settings() are rarely used.
    #
    def outline_settings(visible = 1, symbols_below = 1, symbols_right = 1, auto_style = 0)
      @outline_on    = visible
      @outline_below = symbols_below
      @outline_right = symbols_right
      @outline_style = auto_style

      @outline_changed = 1
    end

    #
    # Deprecated. This is a writeexcel method that is no longer required
    # by WriteXLSX. See below.
    #
    def store_formula(string)
      string.split(/(\$?[A-I]?[A-Z]\$?\d+)/)
    end

    #
    # :call-seq:
    #   write_url(row, column, url [ , format, string, tool_tip ] )
    #
    # Write a hyperlink to a URL in the cell specified by row and column.
    # The hyperlink is comprised of two elements: the visible label and
    # the invisible link. The visible label is the same as the link unless
    # an alternative label is specified. The label parameter is optional.
    # The label is written using the write() method. Therefore it is
    # possible to write strings, numbers or formulas as labels.
    #
    # The hyperlink can be to a http, ftp, mail, internal sheet, or external
    # directory url.
    #
    # write_url methods return:
    #   Returns  0 : normal termination
    #           -1 : insufficient number of arguments
    #           -2 : row or column out of range
    #           -3 : long string truncated to 32767 chars
    #
    # The format parameter is also optional, however, without a format
    # the link won't look like a format.
    #
    # The suggested format is:
    #
    #     format = workbook.add_format(:color => 'blue', :underline => 1)
    #
    # Note, this behaviour is different from writeexcel gem which
    # provides a default hyperlink format if one isn't specified
    # by the user.
    #
    # There are four web style URI's supported:
    # http://, https://, ftp:// and mailto::
    #
    #     worksheet.write_url(0, 0, 'ftp://www.ruby.org/',  format)
    #     worksheet.write_url(1, 0, 'http://www.ruby.com/', format, 'Ruby')
    #     worksheet.write_url('A3', 'http://www.ruby.com/', format)
    #     worksheet.write_url('A4', 'mailto:foo@bar.com', format)
    #
    # There are two local URIs supported: internal: and external:.
    # These are used for hyperlinks to internal worksheet references or
    # external workbook and worksheet references:
    #
    #     worksheet.write_url('A6',  'internal:Sheet2!A1',              format)
    #     worksheet.write_url('A7',  'internal:Sheet2!A1',              format)
    #     worksheet.write_url('A8',  'internal:Sheet2!A1:B2',           format)
    #     worksheet.write_url('A9',  %q{internal:'Sales Data'!A1},      format)
    #     worksheet.write_url('A10', 'external:c:\temp\foo.xlsx',       format)
    #     worksheet.write_url('A11', 'external:c:\foo.xlsx#Sheet2!A1',  format)
    #     worksheet.write_url('A12', 'external:..\foo.xlsx',            format)
    #     worksheet.write_url('A13', 'external:..\foo.xlsx#Sheet2!A1',  format)
    #     worksheet.write_url('A13', 'external:\\\\NET\share\foo.xlsx', format)
    #
    # All of the these URI types are recognised by the write() method, see above.
    #
    # Worksheet references are typically of the form Sheet1!A1. You can
    # also refer to a worksheet range using the standard Excel notation:
    # Sheet1!A1:B2.
    #
    # In external links the workbook and worksheet name must be separated
    # by the # character: external:Workbook.xlsx#Sheet1!A1'.
    #
    # You can also link to a named range in the target worksheet. For
    # example say you have a named range called my_name in the workbook
    # c:\temp\foo.xlsx you could link to it as follows:
    #
    #     worksheet.write_url('A14', 'external:c:\temp\foo.xlsx#my_name')
    #
    # Excel requires that worksheet names containing spaces or non
    # alphanumeric characters are single quoted as follows 'Sales Data'!A1.
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
    # :call-seq:
    #   write_date_time (row, col, date_string [ , format ] )
    #
    # Write a datetime string in ISO8601 "yyyy-mm-ddThh:mm:ss.ss" format as a
    # number representing an Excel date. $format is optional.
    #
    # write_date_time methods return:
    #   Returns  0 : normal termination
    #           -1 : insufficient number of arguments
    #           -2 : row or column out of range
    #           -3 : Invalid date_time, written as string
    #
    # The write_date_time() method can be used to write a date or time
    # to the cell specified by row and column:
    #
    #     worksheet.write_date_time('A1', '2004-05-13T23:20', date_format)
    #
    # The date_string should be in the following format:
    #
    #     yyyy-mm-ddThh:mm:ss.sss
    #
    # This conforms to an ISO8601 date but it should be noted that the
    # full range of ISO8601 formats are not supported.
    #
    # The following variations on the $date_string parameter are permitted:
    #
    #     yyyy-mm-ddThh:mm:ss.sss         # Standard format
    #     yyyy-mm-ddT                     # No time
    #               Thh:mm:ss.sss         # No date
    #     yyyy-mm-ddThh:mm:ss.sssZ        # Additional Z (but not time zones)
    #     yyyy-mm-ddThh:mm:ss             # No fractional seconds
    #     yyyy-mm-ddThh:mm                # No seconds
    #
    # Note that the T is required in all cases.
    #
    # A date should always have a $format, otherwise it will appear
    # as a number, see "DATES AND TIME IN EXCEL" and "CELL FORMATTING".
    # Here is a typical example:
    #
    #     date_format = workbook.add_format(:num_format => 'mm/dd/yy')
    #     worksheet.write_date_time('A1', '2004-05-13T23:20', date_format)
    #
    # Valid dates should be in the range 1900-01-01 to 9999-12-31,
    # for the 1900 epoch and 1904-01-01 to 9999-12-31, for the 1904 epoch.
    # As with Excel, dates outside these ranges will be written as a string.
    #
    # See also the date_time.rb program in the examples directory of the distro.
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
    # :call-seq:
    #   insert_chart(row, column, chart [ , x, y, scale_x, scale_y ] )
    #
    # Insert a chart into a worksheet. The chart argument should be a Chart
    # object or else it is assumed to be a filename of an external binary file.
    # The latter is for backwards compatibility.
    #
    # This method can be used to insert a Chart object into a worksheet.
    # The Chart must be created by the add_chart() Workbook method and
    # it must have the embedded option set.
    #
    #     chart = workbook.add_chart(:type => 'line', :embedded => 1)
    #
    #     # Configure the chart.
    #     ...
    #
    #     # Insert the chart into the a worksheet.
    #     worksheet.insert_chart('E2', chart)
    #
    # See add_chart() for details on how to create the Chart object and
    # Writexlsx::Chart for details on how to configure it. See also the
    # chart_*.rb programs in the examples directory of the distro.
    #
    # The x, y, scale_x and scale_y parameters are optional.
    #
    # The parameters x and y can be used to specify an offset from the top
    # left hand corner of the cell specified by row and column. The offset
    # values are in pixels.
    #
    #     worksheet1.insert_chart('E2', chart, 3, 3)
    #
    # The parameters scale_x and scale_y can be used to scale the inserted
    # image horizontally and vertically:
    #
    #     # Scale the width by 120% and the height by 150%
    #     worksheet.insert_chart('E2', chart, 0, 0, 1.2, 1.5)
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
    # :call-seq:
    #   insert_image(row, column, filename [ , x, y, scale_x, scale_y ] )
    #
    # Partially supported. Currently only works for 96 dpi images. This
    # will be fixed in an upcoming release.
    #--
    # This method can be used to insert a image into a worksheet. The image
    # can be in PNG, JPEG or BMP format. The x, y, scale_x and scale_y
    # parameters are optional.
    #
    #     worksheet1.insert_image('A1', 'ruby.bmp')
    #     worksheet2.insert_image('A1', '../images/ruby.bmp')
    #     worksheet3.insert_image('A1', '.c:\images\ruby.bmp')
    #
    # The parameters x and y can be used to specify an offset from the top
    # left hand corner of the cell specified by row and column. The offset
    # values are in pixels.
    #
    #     worksheet1.insert_image('A1', 'ruby.bmp', 32, 10)
    #
    # The offsets can be greater than the width or height of the underlying
    # cell. This can be occasionally useful if you wish to align two or more
    # images relative to the same cell.
    #
    # The parameters $scale_x and $scale_y can be used to scale the inserted
    # image horizontally and vertically:
    #
    #     # Scale the inserted image: width x 2.0, height x 0.8
    #     worksheet.insert_image('A1', 'perl.bmp', 0, 0, 2, 0.8)
    #
    # See also the images.rb program in the examples directory of the distro.
    #
    # Note: you must call set_row() or set_column() before insert_image()
    # if you wish to change the default dimensions of any of the rows or
    # columns that the image occupies. The height of a row can also change
    # if you use a font that is larger than the default. This in turn will
    # affect the scaling of your image. To avoid this you should explicitly
    # set the height of the row using set_row() if it contains a font size
    # that will change the row height.
    #
    # BMP images must be 24 bit, true colour, bitmaps. In general it is
    # best to avoid BMP images since they aren't compressed.
    #++
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

    #
    # :call-seq:
    #   repeat_formula(row, column, formula [ , format ] )
    #
    # Deprecated. This is a writeexcel gem's method that is no longer
    # required by WriteXLSX.
    #
    # In writeexcel it was computationally expensive to write formulas
    # since they were parsed by a recursive descent parser. The store_formula()
    # and repeat_formula() methods were used as a way of avoiding the overhead
    # of repeated formulas by reusing a pre-parsed formula.
    #
    # In WriteXLSX this is no longer necessary since it is just as quick
    # to write a formula as it is to write a string or a number.
    #
    # The methods remain for backward compatibility but new WriteXLSX
    # programs shouldn't use them.
    #
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
    #            nil if the date is invalid.
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
    # :call-seq:
    #   set_row(row [ , height, format, hidden, level, collapsed ] )
    #
    # This method can be used to change the default properties of a row.
    # All parameters apart from row are optional.
    #
    # The most common use for this method is to change the height of a row:
    #
    #     worksheet.set_row(0, 20)    # Row 1 height set to 20
    #
    # If you wish to set the format without changing the height you can
    # pass nil as the height parameter:
    #
    #     worksheet.set_row(0, nil, format)
    #
    # The format parameter will be applied to any cells in the row that
    # don't have a format. For example
    #
    #     worksheet.set_row(0, nil, format1)      # Set the format for row 1
    #     worksheet.write('A1', 'Hello')          # Defaults to $format1
    #     worksheet.write('B1', 'Hello', format2) # Keeps $format2
    #
    # If you wish to define a row format in this way you should call the
    # method before any calls to write(). Calling it afterwards will overwrite
    # any format that was previously specified.
    #
    # The $hidden parameter should be set to 1 if you wish to hide a row.
    # This can be used, for example, to hide intermediary steps in a
    # complicated calculation:
    #
    #     worksheet.set_row(0, 20,  format, 1)
    #     worksheet.set_row(1, nil, nil,    1)
    #
    # The level parameter is used to set the outline level of the row.
    # Outlines are described in "OUTLINES AND GROUPING IN EXCEL". Adjacent
    # rows with the same outline level are grouped together into a single
    # outline.
    #
    # The following example sets an outline level of 1 for rows 1
    # and 2 (zero-indexed):
    #
    #     worksheet.set_row(1, nil, nil, 0, 1)
    #     worksheet.set_row(2, nil, nil, 0, 1)
    #
    # The hidden parameter can also be used to hide collapsed outlined rows
    # when used in conjunction with the level parameter.
    #
    #     worksheet.set_row(1, nil, nil, 1, 1)
    #     worksheet.set_row(2, nil, nil, 1, 1)
    #
    # For collapsed outlines you should also indicate which row has the
    # collapsed + symbol using the optional collapsed parameter.
    #
    #     worksheet.set_row(3, nil, nil, 0, 0, 1)
    #
    # For a more complete example see the outline.rb and outline_collapsed.rb
    # programs in the examples directory of the distro.
    #
    # Excel allows up to 7 outline levels. Therefore the level parameter
    # should be in the range 0 <= $level <= 7.
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
        write_number(row_first, col_first, *args)
      elsif type == 'blank'
        write_blank(row_first, col_first, *args)
      elsif type == 'date_time'
        write_date_time(row_first, col_first, *args)
      elsif type == 'rich_string'
        write_rich_string(row_first, col_first, *args)
      elsif type == 'url'
        write_url(row_first, col_first, *args)
      elsif type == 'formula'
        write_formula(row_first, col_first, *args)
      elsif type == 'array_formula'
        write_formula_array(row_first, col_first, *args)
      else
        raise "Unknown type '#{type}'"
      end

      # Pad out the rest of the area with formatted blank cells.
      (row_first .. row_last).each do |row|
        (col_first .. col_last).each do |col|
          next if row == row_first && col == col_first
          write_blank(row, col, format)
        end
      end
    end

    #
    # :call-seq:
    #   conditional_formatting(cell_or_cell_range, options)
    #
    # This method handles the interface to Excel conditional formatting.
    #
    # We allow the format to be called on one cell or a range of cells. The
    # hashref contains the formatting parameters and must be the last param:
    #
    #    conditional_formatting(row, col, {...})
    #    conditional_formatting(first_row, first_col, last_row, last_col, {...})
    #
    # conditional_formatting methods return:
    #   Returns  0 : normal termination
    #           -1 : insufficient number of arguments
    #           -2 : row or column out of range
    #           -3 : incorrect parameter.
    #
    # The conditional_format() method is used to add formatting to a cell
    # or range of cells based on user defined criteria.
    #
    #     worksheet.conditional_formatting('A1:J10',
    #         {
    #             :type     => 'cell',
    #             :criteria => '>=',
    #             :value    => 50,
    #             :format   => $format1
    #         }
    #     )
    #
    # This method contains a lot of parameters and is described in detail in
    # a separate section "CONDITIONAL FORMATTING IN EXCEL".
    #
    # See also the conditional_format.pl program in the examples directory of the distro
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
      criteria_type = valid_criteria_type

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

    #
    # :call-seq:
    #   data_validation(cell_or_cell_range, options)
    #
    # The data_validation() method is used to construct an Excel data
    # validation or to limit the user input to a dropdown list of values.
    #
    #     worksheet.data_validation('B3',
    #         {
    #             :validate => 'integer',
    #             :criteria => '>',
    #             :value    => 100
    #         })
    #
    #     worksheet.data_validation('B5:B9',
    #         {
    #             :validate => 'list',
    #             :value    => ['open', 'high', 'close']
    #         })
    #
    # This method contains a lot of parameters and is described in detail
    # in a separate section "DATA VALIDATION IN EXCEL".
    #
    # See also the data_validate.rb program in the examples directory
    # of the distro
    #
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
      criteria_type = valid_criteria_type

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
    # This method is used to hide the gridlines on the screen and printed
    # page. Gridlines are the lines that divide the cells on a worksheet.
    # Screen and printed gridlines are turned on by default in an Excel
    # worksheet. If you have defined your own cell borders you may wish
    # to hide the default gridlines.
    #
    #     worksheet.hide_gridlines
    #
    # The following values of option are valid:
    #
    #     0 : Don't hide gridlines
    #     1 : Hide printed gridlines only
    #     2 : Hide screen and printed gridlines
    #
    # If you don't supply an argument or use nil the default option
    # is true, i.e. only the printed gridlines are hidden.
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

    # Set the option to print the row and column headers on the printed page.
    #
    # An Excel worksheet looks something like the following;
    #
    #      ------------------------------------------
    #     |   |   A   |   B   |   C   |   D   |  ...
    #      ------------------------------------------
    #     | 1 |       |       |       |       |  ...
    #     | 2 |       |       |       |       |  ...
    #     | 3 |       |       |       |       |  ...
    #     | 4 |       |       |       |       |  ...
    #     |...|  ...  |  ...  |  ...  |  ...  |  ...
    #
    # The headers are the letters and numbers at the top and the left of the
    # worksheet. Since these headers serve mainly as a indication of position
    # on the worksheet they generally do not appear on the printed page.
    # If you wish to have them printed you can use the
    # print_row_col_headers() method :
    #
    #     worksheet.print_row_col_headers
    #
    # Do not confuse these headers with page headers as described in the
    # set_header() section above.
    #
    def print_row_col_headers(headers = 1)
      if headers
        @print_headers         = 1
        @print_options_changed = 1
      else
        @print_headers = 0
      end
    end

    #
    # The fit_to_pages() method is used to fit the printed area to a specific
    # number of pages both vertically and horizontally. If the printed area
    # exceeds the specified number of pages it will be scaled down to fit.
    # This guarantees that the printed area will always appear on the
    # specified number of pages even if the page size or margins change.
    #
    #     worksheet1.fit_to_pages(1, 1)    # Fit to 1x1 pages
    #     worksheet2.fit_to_pages(2, 1)    # Fit to 2x1 pages
    #     worksheet3.fit_to_pages(1, 2)    # Fit to 1x2 pages
    #
    # The print area can be defined using the print_area() method
    # as described above.
    #
    # A common requirement is to fit the printed output to n pages wide
    # but have the height be as long as necessary. To achieve this set
    # the height to zero:
    #
    #     worksheet1.fit_to_pages(1, 0)    # 1 page wide and as long as necessary
    #
    # Note that although it is valid to use both fit_to_pages() and
    # set_print_scale() on the same worksheet only one of these options can
    # be active at a time. The last method call made will set the active option.
    #
    # Note that fit_to_pages() will override any manual page breaks that
    # are defined in the worksheet.
    #
    def fit_to_pages(width = 1, height = 1)
      @fit_page           = 1
      @fit_width          = width
      @fit_height         = height
      @page_setup_changed = 1
    end

    #
    # :call-seq:
    #   autofilter(first_row, first_col, last_row, last_col)
    #
    # Set the autofilter area in the worksheet.
    #
    # This method allows an autofilter to be added to a worksheet.
    # An autofilter is a way of adding drop down lists to the headers of a 2D
    # range of worksheet data. This is turn allow users to filter the data
    # based on simple criteria so that some data is shown and some is hidden.
    #
    # To add an autofilter to a worksheet:
    #
    #     worksheet.autofilter(0, 0, 10, 3)
    #     worksheet.autofilter('A1:D11')    # Same as above in A1 notation.
    #
    # Filter conditions can be applied using the filter_column() or
    # filter_column_list() method.
    #
    # See the autofilter.rb program in the examples directory of the distro
    # for a more detailed example.
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
    # The filter_column method can be used to filter columns in a autofilter
    # range based on simple conditions.
    #
    # NOTE: It isn't sufficient to just specify the filter condition.
    # You must also hide any rows that don't match the filter condition.
    # Rows are hidden using the set_row() visible parameter. WriteXLSX cannot
    # do this automatically since it isn't part of the file format.
    # See the autofilter.rb program in the examples directory of the distro
    # for an example.
    #
    # The conditions for the filter are specified using simple expressions:
    #
    #     worksheet.filter_column('A', 'x > 2000')
    #     worksheet.filter_column('B', 'x > 2000 and x < 5000')
    #
    # The column parameter can either be a zero indexed column number or
    # a string column name.
    #
    # The following operators are available:
    #
    #     Operator        Synonyms
    #        ==           =   eq  =~
    #        !=           <>  ne  !=
    #        >
    #        <
    #        >=
    #        <=
    #
    #        and          &&
    #        or           ||
    #
    # The operator synonyms are just syntactic sugar to make you more
    # comfortable using the expressions. It is important to remember that
    # the expressions will be interpreted by Excel and not by ruby.
    #
    # An expression can comprise a single statement or two statements
    # separated by the and and or operators. For example:
    #
    #     'x <  2000'
    #     'x >  2000'
    #     'x == 2000'
    #     'x >  2000 and x <  5000'
    #     'x == 2000 or  x == 5000'
    #
    # Filtering of blank or non-blank data can be achieved by using a value
    # of Blanks or NonBlanks in the expression:
    #
    #     'x == Blanks'
    #     'x == NonBlanks'
    #
    # Excel also allows some simple string matching operations:
    #
    #     'x =~ b*'   # begins with b
    #     'x !~ b*'   # doesn't begin with b
    #     'x =~ *b'   # ends with b
    #     'x !~ *b'   # doesn't end with b
    #     'x =~ *b*'  # contains b
    #     'x !~ *b*'  # doesn't contains b
    #
    # You can also use * to match any character or number and ? to match any
    # single character or number. No other regular expression quantifier is
    # supported by Excel's filters. Excel's regular expression characters can
    # be escaped using ~.
    #
    # The placeholder variable x in the above examples can be replaced by any
    # simple string. The actual placeholder name is ignored internally so the
    # following are all equivalent:
    #
    #     'x     < 2000'
    #     'col   < 2000'
    #     'Price < 2000'
    #
    # Also, note that a filter condition can only be applied to a column
    # in a range specified by the autofilter() Worksheet method.
    #
    # See the autofilter.rb program in the examples directory of the distro
    # for a more detailed example.
    #
    # Note Spreadsheet::WriteExcel supports Top 10 style filters. These aren't
    # currently supported by WriteXLSX but may be added later.
    #
    def filter_column(col, expression)
      raise "Must call autofilter before filter_column" unless @autofilter_area

      col = prepare_filter_column(col)

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
    # Prior to Excel 2007 it was only possible to have either 1 or 2 filter
    # conditions such as the ones shown above in the filter_column method.
    #
    # Excel 2007 introduced a new list style filter where it is possible
    # to specify 1 or more 'or' style criteria. For example if your column
    # contained data for the first six months the initial data would be
    # displayed as all selected as shown on the left. Then if you selected
    # 'March', 'April' and 'May' they would be displayed as shown on the right.
    #
    #     No criteria selected      Some criteria selected.
    #
    #     [/] (Select all)          [X] (Select all)
    #     [/] January               [ ] January
    #     [/] February              [ ] February
    #     [/] March                 [/] March
    #     [/] April                 [/] April
    #     [/] May                   [/] May
    #     [/] June                  [ ] June
    #
    # The filter_column_list() method can be used to represent these types of
    # filters:
    #
    #     worksheet.filter_column_list('A', 'March', 'April', 'May')
    #
    # The column parameter can either be a zero indexed column number or
    # a string column name.
    #
    # One or more criteria can be selected:
    #
    #     worksheet.filter_column_list(0, 'March')
    #     worksheet.filter_column_list(1, 100, 110, 120, 130)
    #
    # NOTE: It isn't sufficient to just specify the filter condition. You must
    # also hide any rows that don't match the filter condition. Rows are hidden
    # using the set_row() visible parameter. WriteXLSX cannot do this
    # automatically since it isn't part of the file format.
    # See the autofilter.rb program in the examples directory of the distro
    # for an example. e conditions for the filter are specified
    # using simple expressions:
    #
    def filter_column_list(col, *tokens)
      tokens.flatten!
      raise "Incorrect number of arguments to filter_column_list" if tokens.empty?
      raise "Must call autofilter before filter_column_list" unless @autofilter_area

      col = prepare_filter_column(col)

      @filter_cols[col] = tokens
      @filter_type[col] = 1           # Default style.
      @filter_on        = 1
    end

    #
    # Store the horizontal page breaks on a worksheet.
    #
    # Add horizontal page breaks to a worksheet. A page break causes all
    # the data that follows it to be printed on the next page. Horizontal
    # page breaks act between rows. To create a page break between rows
    # 20 and 21 you must specify the break at row 21. However in zero index
    # notation this is actually row 20. So you can pretend for a small
    # while that you are using 1 index notation:
    #
    #     worksheet1.set_h_pagebreaks( 20 )    # Break between row 20 and 21
    #
    # The set_h_pagebreaks() method will accept a list of page breaks
    # and you can call it more than once:
    #
    #     worksheet2.set_h_pagebreaks( 20,  40,  60,  80,  100 )    # Add breaks
    #     worksheet2.set_h_pagebreaks( 120, 140, 160, 180, 200 )    # Add some more
    #
    # Note: If you specify the "fit to page" option via the fit_to_pages()
    # method it will override all manual page breaks.
    #
    # There is a silent limitation of about 1000 horizontal page breaks
    # per worksheet in line with an Excel internal limitation.
    #
    def set_h_pagebreaks(*args)
      @hbreaks += args
    end

    #
    # Store the vertical page breaks on a worksheet.
    #
    # Add vertical page breaks to a worksheet. A page break causes all the
    # data that follows it to be printed on the next page. Vertical page breaks
    # act between columns. To create a page break between columns 20 and 21
    # you must specify the break at column 21. However in zero index notation
    # this is actually column 20. So you can pretend for a small while that
    # you are using 1 index notation:
    #
    #     worksheet1.set_v_pagebreaks(20) # Break between column 20 and 21
    #
    # The set_v_pagebreaks() method will accept a list of page breaks
    # and you can call it more than once:
    #
    #     worksheet2.set_v_pagebreaks( 20,  40,  60,  80,  100 )    # Add breaks
    #     worksheet2.set_v_pagebreaks( 120, 140, 160, 180, 200 )    # Add some more
    #
    # Note: If you specify the "fit to page" option via the fit_to_pages()
    # method it will override all manual page breaks.
    #
    def set_v_pagebreaks(*args)
      @vbreaks += args
    end

    #
    # Make any comments in the worksheet visible.
    #
    # This method is used to make all cell comments visible when a worksheet
    # is opened.
    #
    #     worksheet.show_comments
    #
    # Individual comments can be made visible using the visible parameter of
    # the write_comment method (see above):
    #
    #     worksheet.write_comment('C3', 'Hello', :visible => 1)
    #
    # If all of the cell comments have been made visible you can hide
    # individual comments as follows:
    #
    #     worksheet.show_comments
    #     worksheet.write_comment('C3', 'Hello', :visible => 0)
    #
    def show_comments(visible = true)
      @comments_visible = visible
    end

    #
    # Set the default author of the cell comments.
    #
    # This method is used to set the default author of all cell comments.
    #
    #     worksheet.set_comments_author('Ruby')
    #
    # Individual comment authors can be set using the author parameter
    # of the write_comment method.
    #
    # The default comment author is an empty string, '',
    # if no author is specified.
    #
    def set_comments_author(author = '')
      @comments_author = author if author
    end

    def has_comments? # :nodoc:
      !!@has_comments
    end

    def is_chartsheet? # :nodoc:
      !!@is_chartsheet
    end

    #
    # Turn the HoH that stores the comments into an array for easier handling
    # and set the external links.
    #
    def prepare_comments(vml_data_id, vml_shape_id, comment_id) # :nodoc:
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
      (1 .. (count / 1024)).each do |i|
        vml_data_id = "vml_data_id,#{start_data_id + i}"
      end

      @vml_data_id  = vml_data_id
      @vml_shape_id = vml_shape_id

      count
    end

    #
    # Set up chart/drawings.
    #
    def prepare_chart(index, chart_id, drawing_id) # :nodoc:
      drawing_type = 1

      row, col, chart, x_offset, y_offset, scale_x, scale_y  = @charts[index]
      scale_x ||= 0
      scale_y ||= 0

      width  = (0.5 + (480 * scale_x)).to_i
      height = (0.5 + (288 * scale_y)).to_i

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
    # Return nils for data that doesn't exist since Excel can chart series
    # with data missing.
    #
    def get_range_data(row_start, col_start, row_end, col_end)
      # TODO. Check for worksheet limits.

      # Iterate through the table data.
      data = []
      (row_start .. row_end).each do |row_num|
        # Store nil if row doesn't exist.
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
            # Store nil if col doesn't exist.
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
    def extract_filter_tokens(expression = nil) #:nodoc:
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
    def parse_filter_expression(expression, tokens) #:nodoc:
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
    def get_palette_color(index) #:nodoc:
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
    def sort_pagebreaks(*args) #:nodoc:
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
    def position_object_pixels(col_start, row_start, x1, y1, width, height, is_drawing = false) #:nodoc:
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
    def position_object_emus(col_start, row_start, x1, y1, width, height) #:nodoc:
      is_drawing = true
      col_start, row_start, x1, y1, col_end, row_end, x2, y2, x_abs, y_abs =
        position_object_pixels(col_start, row_start, x1, y1, width, height, is_drawing)

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
    def size_col(col) #:nodoc:
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
    def size_row(row) #:nodoc:
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
    def get_shared_string_index(str) #:nodoc:
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
    def prepare_image(index, image_id, drawing_id, width, height, name, image_type) #:nodoc:
      drawing_type = 2
      drawing

      row, col, image, x_offset, y_offset, scale_x, scale_y = @images[index]

      width  *= scale_x
      height *= scale_y

      dimensions = position_object_emus(col, row, x_offset, y_offset, width, height)

      # Convert from pixels to emus.
      width  = int(0.5 + (width * 9_525))
      height = int(0.5 + (height * 9_525))

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
    def comment_params(row, col, string, options = {}) #:nodoc:
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
      params[:width]  = (0.5 + params[:width]).to_i
      params[:height] = (0.5 + params[:height]).to_i

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
    def encode_password(password) #:nodoc:
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
    def write_worksheet #:nodoc:
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
    def write_sheet_pr #:nodoc:
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
    def write_page_set_up_pr #:nodoc:
      return unless fit_page?

      attributes = ['fitToPage', 1]
      @writer.empty_tag('pageSetUpPr', attributes)
    end

    # Write the <dimension> element. This specifies the range of cells in the
    # worksheet. As a special case, empty spreadsheets use 'A1' as a range.
    #
    def write_dimension #:nodoc:
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
    def write_sheet_views #:nodoc:
      @writer.start_tag('sheetViews', [])
      write_sheet_view
      @writer.end_tag('sheetViews')
    end

    def write_sheet_view #:nodoc:
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
    def write_selections #:nodoc:
      @selections.each { |selection| write_selection(*selection) }
    end

    #
    # Write the <selection> element.
    #
    def write_selection(pane, active_cell, sqref) #:nodoc:
      attributes  = []
      (attributes << 'pane' << pane) if pane
      (attributes << 'activeCell' << active_cell) if active_cell
      (attributes << 'sqref' << sqref) if sqref

      @writer.empty_tag('selection', attributes)
    end

    #
    # Write the <sheetFormatPr> element.
    #
    def write_sheet_format_pr #:nodoc:
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
    def write_cols #:nodoc:
      # Exit unless some column have been formatted.
      return if @colinfo.empty?

      @writer.start_tag('cols')
      @colinfo.each {|col_info| write_col_info(*col_info) }

      @writer.end_tag('cols')
    end

    #
    # Write the <col> element.
    #
    def write_col_info(*args) #:nodoc:
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
        width = ((width * max_digit_width + padding) / max_digit_width * 256).to_i/256.0
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
    def write_sheet_data #:nodoc:
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
    def write_rows #:nodoc:
      calculate_spans

      (@dim_rowmin .. @dim_rowmax).each do |row_num|
        # Skip row if it doesn't contain row formatting or cell data.
        next if !@set_rows[row_num] && !@table[row_num] && !@comments[row_num]

        span_index = row_num / 16
        span       = @row_spans[span_index]

        # Write the cells if the row contains data.
        if @table[row_num]
          if !@set_rows[row_num]
            write_row_element(row_num, span)
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
    def write_single_row(current_row = 0) #:nodoc:
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
    def write_row_element(r, spans = nil, height = 15, format = nil, hidden = false, level = 0, collapsed = false, empty_row = false) #:nodoc:
      height    ||= 15
      hidden    ||= 0
      level     ||= 0
      collapsed ||= 0
      empty_row ||= 0
      xf_index = 0

      attributes = ['r',  r + 1]

      xf_index = format.get_xf_index if format

      (attributes << 'spans'        << spans) if spans
      (attributes << 's'            << xf_index) if xf_index != 0
      (attributes << 'customFormat' << 1    ) if format
      (attributes << 'ht'           << height) if height != 15
      (attributes << 'hidden'       << 1    ) if !!hidden && hidden != 0
      (attributes << 'customHeight' << 1    ) if height != 15
      (attributes << 'outlineLevel' << level) if !!level && level != 0
      (attributes << 'collapsed'    << 1    ) if !!collapsed && collapsed != 0

      if empty_row && empty_row != 0
        @writer.empty_tag('row', attributes)
      else
        @writer.start_tag('row', attributes)
      end
    end

    #
    # Write and empty <row> element, i.e., attributes only, no cell data.
    #
    def write_empty_row(*args) #:nodoc:
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
    def write_cell(row, col, cell) #:nodoc:
      type, token, xf = cell

      xf_index = 0
      xf_index = xf.get_xf_index if xf.respond_to?(:get_xf_index)

      range = xl_rowcol_to_cell(row, col)
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
    def write_cell_value(value = '') #:nodoc:
      value ||= ''
      value = value.to_i if value == value.to_i
      @writer.data_element('v', value)
    end

    #
    # Write the cell formula <f> element.
    #
    def write_cell_formula(formula = '') #:nodoc:
      @writer.data_element('f', formula)
    end

    #
    # Write the cell array formula <f> element.
    #
    def write_cell_array_formula(formula, range) #:nodoc:
      attributes = ['t', 'array', 'ref', range]

      @writer.data_element('f', formula, attributes)
    end

    #
    # Write the frozen or split <pane> elements.
    #
    def write_panes #:nodoc:
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
    def write_freeze_panes(row, col, top_row, left_col, type) #:nodoc:
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
      active_pane = set_active_pane_and_cell_selections(row, col, row, col, active_cell, sqref)

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
    def write_split_panes(row, col, top_row, left_col, type) #:nodoc:
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
        top_row  = (0.5 + (y_split - 300) / 20 / 15).to_i
        left_col = (0.5 + (x_split - 390) / 20 / 3 * 4 / 64).to_i
      end

      top_left_cell = xl_rowcol_to_cell(top_row, left_col)

      # If there is no selection set the active cell to the top left cell.
      if !has_selection
        active_cell = top_left_cell
        sqref       = top_left_cell
      end
      active_pane = set_active_pane_and_cell_selections(row, col, top_row, left_col, active_cell, sqref)

      attributes = []
      (attributes << 'xSplit' << x_split) if x_split > 0
      (attributes << 'ySplit' << y_split) if y_split > 0
      attributes << 'topLeftCell' << top_left_cell
      (attributes << 'activePane' << active_pane) if has_selection

      @writer.empty_tag('pane', attributes)
    end

    #
    # Convert column width from user units to pane split width.
    #
    def calculate_x_split_width(width) #:nodoc:
      max_digit_width = 7    # For Calabri 11.
      padding         = 5

      # Convert to pixels.
      if width < 1
        pixels = int(width * 12 + 0.5)
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
    def write_sheet_calc_pr #:nodoc:
      full_calc_on_load = 1

      attributes = ['fullCalcOnLoad', full_calc_on_load]

      @writer.empty_tag('sheetCalcPr', attributes)
    end

    #
    # Write the <phoneticPr> element.
    #
    def write_phonetic_pr #:nodoc:
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
    def write_page_margins #:nodoc:
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
    def write_page_setup #:nodoc:
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
    def write_ext_lst #:nodoc:
      @writer.start_tag('extLst')
      write_ext
      @writer.end_tag('extLst')
    end

    #
    # Write the <ext> element.
    #
    def write_ext #:nodoc:
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
    def write_mx_plv #:nodoc:
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
    def write_merge_cells #:nodoc:
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
    def write_merge_cell(merged_range) #:nodoc:
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
    def write_print_options #:nodoc:
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
    def write_header_footer #:nodoc:
      return unless header_footer_changed?

      @writer.start_tag('headerFooter')
      write_odd_header if @header && @header != ''
      write_odd_footer if @footer && @footer != ''
      @writer.end_tag('headerFooter')
    end

    #
    # Write the <oddHeader> element.
    #
    def write_odd_header #:nodoc:
      @writer.data_element('oddHeader', @header)
    end

    # _write_odd_footer()
    #
    # Write the <oddFooter> element.
    #
    def write_odd_footer #:nodoc:
      @writer.data_element('oddFooter', @footer)
    end

    #
    # Write the <rowBreaks> element.
    #
    def write_row_breaks #:nodoc:
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
    def write_col_breaks #:nodoc:
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
    def write_brk(id, max) #:nodoc:
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
    def write_auto_filter #:nodoc:
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
    def write_autofilters #:nodoc:
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
    def write_filter_column(col_id, type, *filters) #:nodoc:
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
    def write_filters(*filters) #:nodoc:
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
    def write_filter(val) #:nodoc:
      @writer.empty_tag('filter', ['val', val])
    end


    #
    # Write the <customFilters> element.
    #
    def write_custom_filters(*tokens) #:nodoc:
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
    def write_custom_filter(operator, val) #:nodoc:
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
    def write_hyperlinks #:nodoc:
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
    def write_hyperlink_external(row, col, id, location = nil, tooltip = nil) #:nodoc:
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
    def write_hyperlink_internal(row, col, location, display, tooltip = nil) #:nodoc:
      ref = xl_rowcol_to_cell(row, col)

      attributes = ['ref', ref, 'location', location]

      attributes << 'tooltip' << tooltip if tooltip
      attributes << 'display' << display

      @writer.empty_tag('hyperlink', attributes)
    end

    #
    # Write the <tabColor> element.
    #
    def write_tab_color #:nodoc:
      return unless tab_color?

      attributes = ['rgb', get_palette_color(@tab_color)]
      @writer.empty_tag('tabColor', attributes)
    end

    #
    # Write the <sheetProtection> element.
    #
    def write_sheet_protection #:nodoc:
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
    def write_drawings #:nodoc:
      write_drawing(@hlink_count + 1) if drawing?
    end

    #
    # Write the <drawing> element.
    #
    def write_drawing(id) #:nodoc:
      r_id = "rId#{id}"

      attributes = ['r:id', r_id]

      @writer.empty_tag('drawing', attributes)
    end

    #
    # Write the <legacyDrawing> element.
    #
    def write_legacy_drawing #:nodoc:
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
    def write_font(format) #:nodoc:
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
    def write_underline(underline) #:nodoc:
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
    def write_vert_align(val) #:nodoc:
      attributes = ['val', val]

      @rstring.empty_tag('vertAlign', attributes)
    end

    #
    # Write the <color> element.
    #
    def write_color(name, value) #:nodoc:
      attributes = [name, value]

      @rstring.empty_tag('color', attributes)
    end

    #
    # Write the <dataValidations> element.
    #
    def write_data_validations #:nodoc:
      return if @validations.empty?

      attributes = ['count', @validations.size]

      @writer.start_tag('dataValidations', attributes)
      @validations.each { |validation| write_data_validation(validation) }
      @writer.end_tag('dataValidations')
    end

    #
    # Write the <dataValidation> element.
    #
    def write_data_validation(param) #:nodoc:
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
    def write_formula_1(formula) #:nodoc:
      # Convert a list array ref into a comma separated string.
      formula   = %!"#{formula.join(',')}"! if formula.kind_of?(Array)

      formula = formula.sub(/^=/, '') if formula.respond_to?(:sub)

      @writer.data_element('formula1', formula)
    end

    # write_formula_2()
    #
    # Write the <formula2> element.
    #
    def write_formula_2(formula) #:nodoc:
      formula = formula.sub(/^=/, '') if formula.respond_to?(:sub)

      @writer.data_element('formula2', formula)
    end

    # in Perl module : _write_formula()
    #
    def write_formula_tag(data) #:nodoc:
      @writer.data_element('formula', data)
    end

    #
    # Write the Worksheet conditional formats.
    #
    def write_conditional_formats #:nodoc:
      ranges = @cond_formats.keys.sort
      return if ranges.empty?

      ranges.each { |range| write_conditional_formatting(range, @cond_formats[range]) }
    end

    #
    # Write the <conditionalFormatting> element.
    #
    def write_conditional_formatting(range, params) #:nodoc:
      attributes = ['sqref', range]

      @writer.start_tag('conditionalFormatting', attributes)

      params.each { |param| write_cf_rule(param) }

      @writer.end_tag('conditionalFormatting')
    end

    #
    # Write the <cfRule> element.
    #
    def write_cf_rule(param) #:nodoc:
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

    def store_data_to_table(row, col, data) #:nodoc:
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
    def calculate_spans #:nodoc:
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

    def xf(format) #:nodoc:
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
    def shared_string_index(str) #:nodoc:
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
    def convert_name_area(row_num_1, col_num_1, row_num_2, col_num_2) #:nodoc:
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
    def quote_sheetname(sheetname) #:nodoc:
      return sheetname if sheetname =~ /^Sheet\d+$/
      return "'#{sheetname}'"
    end

    def fit_page? #:nodoc:
      if @fit_page
        @fit_page != 0
      else
        false
      end
    end

    def filter_on? #:nodoc:
      if @filter_on
        @filter_on != 0
      else
        false
      end
    end

    def tab_color? #:nodoc:
      if @tab_color
        @tab_color != 0
      else
        false
      end
    end

    def zoom_scale_normal? #:nodoc:
      !!@zoom_scale_normal
    end

    def page_view? #:nodoc:
      !!@page_view
    end

    def right_to_left? #:nodoc:
      !!@right_to_left
    end

    def show_zeros? #:nodoc:
      !!@show_zeros
    end

    def screen_gridlines? #:nodoc:
      !!@screen_gridlines
    end

    def protect? #:nodoc:
      !!@protect
    end

    def autofilter_ref? #:nodoc:
      !!@autofilter_ref
    end

    def date_1904? #:nodoc:
      @workbook.date_1904?
    end

    def print_options_changed? #:nodoc:
      !!@print_options_changed
    end

    def hcenter? #:nodoc:
      !!@hcenter
    end

    def vcenter? #:nodoc:
      !!@vcenter
    end

    def print_headers? #:nodoc:
      !!@print_headers
    end

    def print_gridlines? #:nodoc:
      !!@print_gridlines
    end

    def page_setup_changed? #:nodoc:
      !!@page_setup_changed
    end

    def orientation? #:nodoc:
      !!@orientation
    end

    def header_footer_changed? #:nodoc:
      !!@header_footer_changed
    end

    def drawing? #:nodoc:
      !!@drawing
    end

    def remove_white_space(margin) #:nodoc:
      if margin.respond_to?(:gsub)
        margin.gsub(/[^\d\.]/, '')
      else
        margin
      end
    end

    # List of valid criteria types.
    def valid_criteria_type  # :nodoc:
      {
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
    end

    def set_active_pane_and_cell_selections(row, col, top_row, left_col, active_cell, sqref) # :nodoc:
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
      active_pane
    end

    def prepare_filter_column(col) # :nodoc:
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
      col
    end
  end
end
