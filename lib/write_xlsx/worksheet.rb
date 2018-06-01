# -*- coding: utf-8 -*-
require 'write_xlsx/package/xml_writer_simple'
require 'write_xlsx/package/button'
require 'write_xlsx/colors'
require 'write_xlsx/format'
require 'write_xlsx/drawing'
require 'write_xlsx/sparkline'
require 'write_xlsx/compatibility'
require 'write_xlsx/utility'
require 'write_xlsx/package/conditional_format'
require 'write_xlsx/worksheet/cell_data'
require 'write_xlsx/worksheet/data_validation'
require 'write_xlsx/worksheet/hyperlink'
require 'write_xlsx/worksheet/page_setup'
require 'tempfile'

module Writexlsx
  #
  # A new worksheet is created by calling the add_worksheet() method from a
  # workbook object:
  #
  #     worksheet1 = workbook.add_worksheet
  #     worksheet2 = workbook.add_worksheet
  #
  # The following methods are available through a new worksheet:
  #
  # * {#write}[#method-i-write]
  # * write_number
  # * write_string
  # * write_rich_string
  # * write_blank
  # * write_row
  # * write_col
  # * write_date_time
  # * write_url
  # * write_formula
  # * write_comment
  # * show_comments
  # * {#comments_author=}[#method-i-comments_author-3D]
  # * insert_image
  # * insert_chart
  # * insert_shape
  # * insert_button
  # * data_validation
  # * conditional_formatting
  # * add_sparkline
  # * add_table
  # * {#name}[#method-i-name]
  # * {#activate}[#method-i-activate]
  # * {#select}[#method-i-select]
  # * {#hide}[#method-i-hide]
  # * set_first_sheet
  # * {#protect}[#method-i-protect]
  # * set_selection
  # * set_row
  # * set_column
  # * outline_settings
  # * freeze_panes
  # * split_panes
  # * merge_range
  # * merge_range_type
  # * {#zoom=}[#method-i-zoom-3D]
  # * right_to_left
  # * hide_zero
  # * {#tab_color=}[#method-i-tab_color-3D]
  # * {#autofilter}[#method-i-autofilter]
  # * filter_column
  # * filter_column_list
  #
  # == PAGE SET-UP METHODS
  #
  # Page set-up methods affect the way that a worksheet looks
  # when it is printed. They control features such as page headers and footers
  # and margins. These methods are really just standard worksheet methods.
  #
  # The following methods are available for page set-up:
  #
  # * set_landscape
  # * set_portrait
  # * set_page_view
  # * {paper=}[#method-i-paper-3D]
  # * center_horizontally
  # * center_vertically
  # * {margins=}[#method-i-margin-3D]
  # * set_header
  # * set_footer
  # * repeat_rows
  # * repeat_columns
  # * hide_gridlines
  # * print_row_col_headers
  # * print_area
  # * print_across
  # * fit_to_pages
  # * {start_page=}[#method-i-start_page-3D]
  # * {print_scale=}[#method-i-print_scale-3D]
  # * set_h_pagebreaks
  # * set_v_pagebreaks
  #
  # A common requirement when working with WriteXLSX is to apply the same
  # page set-up features to all of the worksheets in a workbook. To do this
  # you can use the sheets() method of the workbook class to access the array
  # of worksheets in a workbook:
  #
  #   workbook.sheets.each do |worksheet|
  #     worksheet.set_landscape
  #   end
  #
  # ==Cell notation
  #
  # WriteXLSX supports two forms of notation to designate the position of cells:
  # Row-column notation and A1 notation.
  #
  # Row-column notation uses a zero based index for both row and column
  # while A1 notation uses the standard Excel alphanumeric sequence of column
  # letter and 1-based row. For example:
  #
  #     (0, 0)      # The top left cell in row-column notation.
  #     ('A1')      # The top left cell in A1 notation.
  #
  #     (1999, 29)  # Row-column notation.
  #     ('AD2000')  # The same cell in A1 notation.
  #
  # Row-column notation is useful if you are referring to cells
  # programmatically:
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
  # Note: in Excel it is also possible to use a R1C1 notation. This is not
  # supported by WriteXLSX.
  #
  # == FORMULAS AND FUNCTIONS IN EXCEL
  #
  # === Introduction
  #
  # The following is a brief introduction to formulas and functions in Excel
  # and WriteXLSX.
  #
  # A formula is a string that begins with an equals sign:
  #
  #     '=A1+B1'
  #     '=AVERAGE(1, 2, 3)'
  #
  # The formula can contain numbers, strings, boolean values, cell references,
  # cell ranges and functions. Named ranges are not supported. Formulas should
  # be written as they appear in Excel, that is cells and functions must be
  # in uppercase.
  #
  # Cells in Excel are referenced using the A1 notation system where the column
  # is designated by a letter and the row by a number. Columns range from +A+
  # to +XFD+ i.e. 0 to 16384, rows range from 1 to 1048576.
  # The Writexlsx::Utility module that is included in the distro contains
  # helper functions for dealing with A1 notation, for example:
  #
  #     require 'write_xlsx'
  #
  #     include Writexlsx::Utility
  #
  #     row, col = xl_cell_to_rowcol('C2')    # (1, 2)
  #     str      = xl_rowcol_to_cell(1, 2)    # C2
  #
  # The Excel +$+ notation in cell references is also supported. This allows
  # you to specify whether a row or column is relative or absolute. This only
  # has an effect if the cell is copied. The following examples show relative
  # and absolute values.
  #
  #     '=A1'   # Column and row are relative
  #     '=$A1'  # Column is absolute and row is relative
  #     '=A$1'  # Column is relative and row is absolute
  #     '=$A$1' # Column and row are absolute
  #
  # Formulas can also refer to cells in other worksheets of the current
  # workbook. For example:
  #
  #     '=Sheet2!A1'
  #     '=Sheet2!A1:A5'
  #     '=Sheet2:Sheet3!A1'
  #     '=Sheet2:Sheet3!A1:A5'
  #     %Q{='Test Data'!A1}
  #     %Q{='Test Data1:Test Data2'!A1}
  #
  # The sheet reference and the cell reference are separated by +!+ the
  # exclamation mark symbol. If worksheet names contain spaces, commas or
  # parentheses then Excel requires that the name is enclosed in single
  # quotes as shown in the last two examples above. In order to avoid using
  # a lot of escape characters you can use the quote operator +%Q{}+ to
  # protect the quotes. Only valid sheet names that have been added using the
  # add_worksheet() method can be used in formulas. You cannot reference
  # external workbooks.
  #
  # The following table lists the operators that are available in Excel's
  # formulas. The majority of the operators are the same as Ruby's,
  # differences are indicated:
  #
  #     Arithmetic operators:
  #     =====================
  #     Operator  Meaning                   Example
  #        +      Addition                  1+2
  #        -      Subtraction               2-1
  #        *      Multiplication            2*3
  #        /      Division                  1/4
  #        ^      Exponentiation            2^3      # Equivalent to **
  #        -      Unary minus               -(1+2)
  #        %      Percent (Not modulus)     13%
  #
  #
  #     Comparison operators:
  #     =====================
  #     Operator  Meaning                   Example
  #         =     Equal to                  A1 =  B1 # Equivalent to ==
  #         <>    Not equal to              A1 <> B1 # Equivalent to !=
  #         >     Greater than              A1 >  B1
  #         <     Less than                 A1 <  B1
  #         >=    Greater than or equal to  A1 >= B1
  #         <=    Less than or equal to     A1 <= B1
  #
  #
  #     String operator:
  #     ================
  #     Operator  Meaning                   Example
  #         &     Concatenation             "Hello " & "World!" # [1]
  #
  #
  #     Reference operators:
  #     ====================
  #     Operator  Meaning                   Example
  #         :     Range operator            A1:A4               # [2]
  #         ,     Union operator            SUM(1, 2+2, B3)     # [3]
  #
  #
  #     Notes:
  #     [1]: Equivalent to "Hello " + "World!" in Ruby.
  #     [2]: This range is equivalent to cells A1, A2, A3 and A4.
  #     [3]: The comma behaves like the list separator in Perl.
  #
  # The range and comma operators can have different symbols in non-English
  # versions of Excel. These may be supported in a later version of WriteXLSX.
  # In the meantime European users of Excel take note:
  #
  #     worksheet.write('A1', '=SUM(1; 2; 3)')   # Wrong!!
  #     worksheet.write('A1', '=SUM(1, 2, 3)')   # Okay
  #
  # For a general introduction to Excel's formulas and an explanation of the
  # syntax of the function refer to the Excel help files or the following:
  # http://office.microsoft.com/en-us/assistance/CH062528031033.aspx.
  #
  # If your formula doesn't work in Excel::Writer::XLSX try the following:
  #
  #     1. Verify that the formula works in Excel.
  #     2. Ensure that cell references and formula names are in uppercase.
  #     3. Ensure that you are using ':' as the range operator, A1:A4.
  #     4. Ensure that you are using ',' as the union operator, SUM(1,2,3).
  #     5. If you verify that the formula works in Gnumeric, OpenOffice.org
  #        or LibreOffice, make sure to note items 2-4 above, since these
  #        applications are more flexible than Excel with formula syntax.
  #
  class Worksheet
    include Writexlsx::Utility

    MAX_DIGIT_WIDTH = 7    # For Calabri 11.  # :nodoc:
    PADDING         = 5                       # :nodoc:

    attr_reader :index # :nodoc:
    attr_reader :charts, :images, :tables, :shapes, :drawing # :nodoc:
    attr_reader :header_images, :footer_images # :nodoc:
    attr_reader :vml_drawing_links # :nodoc:
    attr_reader :vml_data_id # :nodoc:
    attr_reader :vml_header_id # :nodoc:
    attr_reader :autofilter_area # :nodoc:
    attr_reader :writer, :set_rows, :col_formats # :nodoc:
    attr_reader :vml_shape_id # :nodoc:
    attr_reader :comments, :comments_author # :nodoc:
    attr_accessor :dxf_priority # :nodoc:
    attr_reader :vba_codename # :nodoc:

    def initialize(workbook, index, name) #:nodoc:
      @writer = Package::XMLWriterSimple.new

      @workbook = workbook
      @index = index
      @name = name
      @colinfo = {}
      @cell_data_table = {}
      @excel_version = 2007
      @palette = workbook.palette

      @page_setup = PageSetup.new

      @screen_gridlines = true
      @show_zeros = true
      @dim_rowmin = nil
      @dim_rowmax = nil
      @dim_colmin = nil
      @dim_colmax = nil
      @selections = []
      @panes = []

      @tab_color  = 0

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

      @last_shape_id          = 1
      @rel_count              = 0
      @hlink_count            = 0
      @external_hyper_links   = []
      @external_drawing_links = []
      @external_comment_links = []
      @external_vml_links     = []
      @external_table_links   = []
      @drawing_links          = []
      @vml_drawing_links      = []
      @charts                 = []
      @images                 = []
      @tables                 = []
      @sparklines             = []
      @shapes                 = []
      @shape_hash             = {}
      @header_images          = []
      @footer_images          = []

      @outline_row_level = 0
      @outline_col_level = 0

      @original_row_height    = 15
      @default_row_height     = 15
      @default_row_pixels     = 20
      @default_col_pixels     = 64
      @default_row_rezoed     = 0

      @merge = []

      @has_vml        = false
      @has_header_vml = false
      @comments = Package::Comments.new(self)
      @buttons_array          = []
      @header_images_array    = []

      @validations = []

      @cond_formats = {}
      @dxf_priority = 1

      if excel2003_style?
        @original_row_height      = 12.75
        @default_row_height       = 12.75
        @default_row_pixels       = 17
        self::margins_left_right  = 0.75
        self::margins_top_bottom  = 1
        @page_setup.margin_header = 0.5
        @page_setup.margin_footer = 0.5
        @page_setup.header_footer_aligns = false
      end
    end

    def set_xml_writer(filename) #:nodoc:
      @writer.set_xml_writer(filename)
    end

    def assemble_xml_file #:nodoc:
      write_xml_declaration do
        @writer.tag_elements('worksheet', write_worksheet_attributes) do
          write_sheet_pr
          write_dimension
          write_sheet_views
          write_sheet_format_pr
          write_cols
          write_sheet_data
          write_sheet_protection
          # write_sheet_calc_pr
          write_phonetic_pr if excel2003_style?
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
          write_legacy_drawing_hf
          write_table_parts
          write_ext_sparklines
        end
      end
    end

    #
    # The name method is used to retrieve the name of a worksheet.
    # For example:
    #
    #     workbook.sheets.each do |sheet|
    #       print sheet.name
    #     end
    #
    # For reasons related to the design of WriteXLSX and to the internals
    # of Excel there is no set_name() method. The only way to set the
    # worksheet name is via the Workbook#add_worksheet() method.
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
    # can be selected via the select() method, however only one
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
      @workbook.activesheet = 0 if @workbook.activesheet == @index
      @workbook.firstsheet  = 0 if @workbook.firstsheet  == @index
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
      @workbook.firstsheet = @index
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
    # for an illustrative example and the +set_locked+ and +set_hidden+ format
    # methods in "CELL FORMATTING", see Format.
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
    # by passing a hash with any or all of the following keys:
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
    #
    # The default boolean values are shown above. Individual elements
    # can be protected as follows:
    #
    #     worksheet.protect('drowssap', { :insert_rows => true } )
    #
    def protect(password = nil, options = {})
      check_parameter(options, protect_default_settings.keys, 'protect')
      @protect = protect_default_settings.merge(options)

      # Set the password after the user defined values.
      @protect[:password] =
        sprintf("%X", encode_password(password)) if password && password != ''
    end

    def protect_default_settings  # :nodoc:
      {
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
        :select_unlocked_cells => true
      }
    end
    private :protect_default_settings

    #
    # :call-seq:
    #   set_column(firstcol, lastcol, width, format, hidden, level, collapsed)
    #
    # This method can be used to change the default properties of a single
    # column or a range of columns. All parameters apart from +first_col+
    # and +last_col+ are optional.
    #
    # If set_column() is applied to a single column the value of +first_col+
    # and +last_col+ should be the same. In the case where +last_col+ is zero
    # it is set to the same value as +first_col+.
    #
    # It is also possible, and generally clearer, to specify a column range
    # using the form of A1 notation used for columns. See the note about
    # {"Cell notation"}[#label-Cell+notation].
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
    # See {"CELL FORMATTING"}[Format.html#label-CELL+FORMATTING].
    # If you wish to set the format without changing the width you can pass
    # nil as the width parameter:
    #
    #     worksheet.set_column(0, 0, nil, format)
    #
    # The format parameter will be applied to any cells in the column that
    # don't have a format. For example
    #
    #     worksheet.set_column('A:A', nil, format1)    # Set format for col 1
    #     worksheet.write('A1', 'Hello')               # Defaults to format1
    #     worksheet.write('A2', 'Hello', format2)      # Keeps format2
    #
    # If you wish to define a column format in this way you should call the
    # method before any calls to {#write()}[#method-i-write].
    # If you call it afterwards it won't have any effect.
    #
    # A default row format takes precedence over a default column format
    #
    #     worksheet.set_row( 0, nil, format1 )           # Set format for row 1
    #     worksheet.set_column( 'A:A', nil, format2 )    # Set format for col 1
    #     worksheet.write( 'A1', 'Hello' )               # Defaults to format1
    #     worksheet.write( 'A2', 'Hello' )               # Defaults to format2
    #
    # The +hidden+ parameter should be set to 1 if you wish to hide a column.
    # This can be used, for example, to hide intermediary steps in a
    # complicated calculation:
    #
    #     worksheet.set_column( 'D:D', 20,  format, 1 )
    #     worksheet.set_column( 'E:E', nil, nil,    1 )
    #
    # The +level+ parameter is used to set the outline level of the column.
    # Outlines are described in
    # {"OUTLINES AND GROUPING IN EXCEL"}["method-i-set_row-label-OUTLINES+AND+GROUPING+IN+EXCEL"].
    # Adjacent columns with the same outline level are grouped together into
    # a single outline.
    #
    # The following example sets an outline level of 1 for columns B to G:
    #
    #     worksheet.set_column( 'B:G', nil, nil, 0, 1 )
    #
    # The +hidden+ parameter can also be used to hide collapsed outlined
    # columns when used in conjunction with the +level+ parameter.
    #
    #     worksheet.set_column( 'B:G', nil, nil, 1, 1 )
    #
    # For collapsed outlines you should also indicate which row has the
    # collapsed + symbol using the optional +collapsed+ parameter.
    #
    #     worksheet.set_column( 'H:H', nil, nil, 0, 0, 1 )
    #
    # For a more complete example see the outline.rb and outline_collapsed.rb
    # programs in the examples directory of the distro.
    #
    # Excel allows up to 7 outline levels. Therefore the level parameter
    # should be in the range <tt>0 <= level <= 7</tt>.
    #
    def set_column(*args)
      # Check for a cell reference in A1 notation and substitute row and column
      if args[0] =~ /^\D/
        row1, firstcol, row2, lastcol, *data = substitute_cellref(*args)
      else
        firstcol, lastcol, *data = args
      end

      # Ensure at least firstcol, lastcol and width
      return unless firstcol && lastcol && !data.empty?

      # Assume second column is the same as first if 0. Avoids KB918419 bug.
      lastcol = firstcol unless ptrue?(lastcol)

      # Ensure 2nd col is larger than first. Also for KB918419 bug.
      firstcol, lastcol = lastcol, firstcol if firstcol > lastcol

      width, format, hidden, level, collapsed = data

      # Check that cols are valid and store max and min values with default row.
      # NOTE: The check shouldn't modify the row dimensions and should only modify
      #       the column dimensions in certain cases.
      ignore_row = 1
      ignore_col = 1
      ignore_col = 0 if format.respond_to?(:xf_index)   # Column has a format.
      ignore_col = 0 if width && ptrue?(hidden)         # Column has a width but is hidden

      check_dimensions_and_update_max_min_values(0, firstcol, ignore_row, ignore_col)
      check_dimensions_and_update_max_min_values(0, lastcol,  ignore_row, ignore_col)

      # Set the limits for the outline levels (0 <= x <= 7).
      level ||= 0
      level = 0 if level < 0
      level = 7 if level > 7

      @outline_col_level = level if level > @outline_col_level

      # Store the column data based on the first column. Padded for sorting.
      @colinfo[sprintf("%05d", firstcol)] = [firstcol, lastcol, width, format, hidden, level, collapsed]

      # Store the column change to allow optimisations.
      @col_size_changed = 1

      # Store the col sizes for use when calculating image vertices taking
      # hidden columns into account. Also store the column formats.
      width = 0 if ptrue?(hidden)         # Set width to zero if hidden

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
    # in which case +last_row+ and +last_col+ can be omitted. The active cell
    # within a selected range is determined by the order in which +first+ and
    # +last+ are specified. It is also possible to specify a cell or a range
    # using A1 notation. See the note about
    # {"Cell notation"}[#label-Cell+notation].
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

      row_first, col_first, row_last, col_last = row_col_notation(args)
      active_cell = xl_rowcol_to_cell(row_first, col_first)

      if row_last  # Range selection.
        # Swap last row/col for first row/col as necessary
        row_first, row_last = row_last, row_first if row_first > row_last
        col_first, col_last = col_last, col_first if col_first > col_last

        # If the first and last cell are the same write a single cell.
        if row_first == row_last && col_first == col_last
          sqref = active_cell
        else
          sqref = xl_range(row_first, row_last, col_first, col_last)
        end
      else          # Single cell selection.
        sqref = active_cell
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
    # <tt>Window->Freeze</tt> Panes menu command in Excel
    #
    # The parameters +row+ and +col+ are used to specify the location of
    # the split. It should be noted that the split is specified at the
    # top or left of a cell and that the method uses zero based indexing.
    # Therefore to freeze the first row of a worksheet it is necessary
    # to specify the split at row 2 (which is 1 as the zero-based index).
    # This might lead you to think that you are using a 1 based index
    # but this is not the case.
    #
    # You can set one of the row and +col+ parameters as zero if you
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
    # The parameters +top_row+ and +left_col+ are optional. They are used
    # to specify the top-most or left-most visible row or column in the
    # scrolling region of the panes. For example to freeze the first row
    # and to have the scrolling region begin at row twenty:
    #
    #     worksheet.freeze_panes(1, 0, 20, 0)
    #
    # You cannot use A1 notation for the +top_row+ and +left_col+ parameters.
    #
    # See also the panes.rb program in the examples directory of the
    # distribution.
    #
    def freeze_panes(*args)
      return if args.empty?

      # Check for a cell reference in A1 notation and substitute row and column.
      row, col, top_row, left_col, type = row_col_notation(args)

      col      ||= 0
      top_row  ||= row
      left_col ||= col
      type     ||= 0

      @panes   = [row, col, top_row, left_col, type ]
    end

    #
    # :call-seq:
    #   split_panes(y, x, top_row, left_col)
    #
    # Set panes and mark them as split.
    #--
    # Implementers note. The API for this method doesn't map well from the XLS
    # file format and isn't sufficient to describe all cases of split panes.
    # It should probably be something like:
    #
    #     split_panes(y, x, top_row, left_col, offset_row, offset_col)
    #
    # I'll look at changing this if it becomes an issue.
    #++
    # This method can be used to divide a worksheet into horizontal or vertical
    # regions known as panes. This method is different from the freeze_panes()
    # method in that the splits between the panes will be visible to the user
    # and each pane will have its own scroll bars.
    #
    # The parameters +y+ and +x+ are used to specify the vertical and horizontal
    # position of the split. The units for y and x are the same as those
    # used by Excel to specify row height and column width. However, the
    # vertical and horizontal units are different from each other. Therefore
    # you must specify the y and x parameters in terms of the row heights
    # and column widths that you have set or the default values which are 15
    # for a row and 8.43 for a column.
    #
    # You can set one of the +y+ and +x+ parameters as zero if you do not want
    # either a vertical or horizontal split. The parameters +top_row+ and
    # +left_col+ are optional. They are used to specify the top-most or
    # left-most visible row or column in the bottom-right pane.
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
      @page_setup.orientation        = true
      @page_setup.page_setup_changed = true
    end

    #
    # Set the page orientation as landscape.
    #
    def set_landscape
      @page_setup.orientation         = false
      @page_setup.page_setup_changed  = true
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
    # The tab_color=() method is used to change the colour of the worksheet
    # tab. This feature is only available in Excel 2002 and later. You can use
    # one of the standard colour names provided by the Format object or a
    # colour index.
    # See "COLOURS IN EXCEL" and the set_custom_color() method.
    #
    #     worksheet1.tab_color = 'red'
    #     worksheet2.tab_color = 0x0C
    #
    # See the tab_colors.rb program in the examples directory of the distro.
    #
    def tab_color=(color)
      @tab_color = Colors.new.color(color)
    end

    # This method is deprecated. use tab_color=().
    def set_tab_color(color)
      put_deprecate_message("#{self}.set_tab_color")
      self.tab_color = color
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
    def paper=(paper_size)
      @page_setup.paper = paper_size
    end

    def set_paper(paper_size)
      put_deprecate_message("#{self}.set_paper")
      self::paper = paper_size
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
    def set_header(string = '', margin = 0.3, options = {})
      raise 'Header string must be less than 255 characters' if string.length >= 255
      # Replace the Excel placeholder &[Picture] with the internal &G.
      @page_setup.header = string.gsub(/&\[Picture\]/, '&G')

      if string.size >= 255
        raise 'Header string must be less than 255 characters'
      end

      if options[:align_with_margins]
        @page_setup.header_footer_aligns = options[:align_with_margins]
      end

      if options[:scale_with_doc]
        @page_setup.header_footer_scales = options[:scale_with_doc]
      end

      # Reset the array in case the function is called more than once.
      @header_images = []

      [
       [:image_left, 'LH'], [:image_center, 'CH'], [:image_right, 'RH']
      ].each do |p|
        if options[p.first]
          @header_images << [options[p.first], p.last]
        end
      end

      # placeholeder /&G/ の数
      placeholder_count = @page_setup.header.scan(/&G/).count

      image_count = @header_images.count

      if image_count != placeholder_count
        raise "Number of header image (#{image_count}) doesn't match placeholder count (#{placeholder_count}) in string: #{@page_setup.header}"
      end

      @has_header_vml = true if image_count > 0

      @page_setup.margin_header         = margin || 0.3
      @page_setup.header_footer_changed = true
    end

    #
    # Set the page footer caption and optional margin.
    #
    # The syntax of the set_footer() method is the same as set_header()
    #
    def set_footer(string = '', margin = 0.3, options = {})
      raise 'Footer string must be less than 255 characters' if string.length >= 255

      @page_setup.footer                = string.dup

      # Replace the Excel placeholder &[Picture] with the internal &G.
      @page_setup.footer = string.gsub(/&\[Picture\]/, '&G')

      if string.size >= 255
        raise 'Header string must be less than 255 characters'
      end

      if options[:align_with_margins]
        @page_setup.header_footer_aligns = options[:align_with_margins]
      end

      if options[:scale_with_doc]
        @page_setup.header_footer_scales = options[:scale_with_doc]
      end

      # Reset the array in case the function is called more than once.
      @footer_images = []

      [
       [:image_left, 'LF'], [:image_center, 'CF'], [:image_right, 'RF']
      ].each do |p|
        if options[p.first]
          @footer_images << [options[p.first], p.last]
        end
      end

      # placeholeder /&G/ の数
      placeholder_count = @page_setup.footer.scan(/&G/).count

      image_count = @footer_images.count

      if image_count != placeholder_count
        raise "Number of footer image (#{image_count}) doesn't match placeholder count (#{placeholder_count}) in string: #{@page_setup.footer}"
      end

      @has_header_vml = true if image_count > 0

      @page_setup.margin_footer         = margin
      @page_setup.header_footer_changed = true
    end

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
    # There are several methods available for setting the worksheet margins
    # on the printed page:
    #
    #     margins=()                # Set all margins to the same value
    #     margins_left_right=()     # Set left and right margins to the same value
    #     margins_top_bottom=()     # Set top and bottom margins to the same value
    #     margin_left=()            # Set left margin
    #     margin_right=()           # Set right margin
    #     margin_top=()             # Set top margin
    #     margin_bottom=()          # Set bottom margin
    #
    # All of these methods take a distance in inches as a parameter.
    # Note: 1 inch = 25.4mm. ;-) The default left and right margin is 0.7 inch.
    # The default top and bottom margin is 0.75 inch. Note, these defaults
    # are different from the defaults used in the binary file format
    # by writeexcel gem.
    #
    def margins=(margin)
      self::margin_left   = margin
      self::margin_right  = margin
      self::margin_top    = margin
      self::margin_bottom = margin
    end

    #
    # Set the left and right margins to the same value in inches.
    # See set_margins
    #
    def margins_left_right=(margin)
      self::margin_left  = margin
      self::margin_right = margin
    end

    #
    # Set the top and bottom margins to the same value in inches.
    # See set_margins
    #
    def margins_top_bottom=(margin)
      self::margin_top    = margin
      self::margin_bottom = margin
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

    #
    # set_margin_* methods are deprecated. use margin_*=().
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
      put_deprecate_message("#{self}.set_margins")
      self::margins = margin
    end

    #
    # this method is deprecated. use margin_left_right=().
    # Set the left and right margins to the same value in inches.
    # See set_margins
    #
    def set_margins_LR(margin)
      put_deprecate_message("#{self}.set_margins_LR")
      self::margins_left_right = margin
    end

    #
    # this method is deprecated. use margin_top_bottom=().
    # Set the top and bottom margins to the same value in inches.
    # See set_margins
    #
    def set_margins_TB(margin)
      put_deprecate_message("#{self}.set_margins_TB")
      self::margins_top_bottom = margin
    end

    #
    # this method is deprecated. use margin_left=()
    # Set the left margin in inches.
    # See set_margins
    #
    def set_margin_left(margin = 0.7)
      put_deprecate_message("#{self}.set_margin_left")
      self::margin_left = margin
    end

    #
    # this method is deprecated. use margin_right=()
    # Set the right margin in inches.
    # See set_margins
    #
    def set_margin_right(margin = 0.7)
      put_deprecate_message("#{self}.set_margin_right")
      self::margin_right = margin
    end

    #
    # this method is deprecated. use margin_top=()
    # Set the top margin in inches.
    # See set_margins
    #
    def set_margin_top(margin = 0.75)
      put_deprecate_message("#{self}.set_margin_top")
      self::margin_top = margin
    end

    #
    # this method is deprecated. use margin_bottom=()
    # Set the bottom margin in inches.
    # See set_margins
    #
    def set_margin_bottom(margin = 0.75)
      put_deprecate_message("#{self}.set_margin_bottom")
      self::margin_bottom = margin
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
    # For large Excel documents it is often desirable to have the first
    # column or columns of the worksheet print out at the left hand side
    # of each page. This can be achieved by using the repeat_columns()
    # method. The parameters first_column and last_column are zero based.
    # The last_column parameter is optional if you only wish to specify
    # one column. You can also specify the columns using A1 column
    # notation, see the note about {"Cell notation"}[#label-Cell+notation].
    #
    #     worksheet1.repeat_columns(0)        # Repeat the first column
    #     worksheet2.repeat_columns(0, 1)     # Repeat the first two columns
    #     worksheet3.repeat_columns('A:A')    # Repeat the first column
    #     worksheet4.repeat_columns('A:B')    # Repeat the first two columns
    #
    def repeat_columns(*args)
      if args[0] =~ /^\D/
        dummy, first_col, dummy, last_col = substitute_cellref(*args)
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
    # A1 notation, see the note about {"Cell notation"}[#label-Cell+notation].
    #
    #     worksheet1.print_area( 'A1:H20' );    # Cells A1 to H20
    #     worksheet2.print_area( 0, 0, 19, 7 ); # The same
    #     worksheet2.print_area( 'A:H' );       # Columns A to H if rows have data
    #
    def print_area(*args)
      return @page_setup.print_area.dup if args.empty?
      row1, col1, row2, col2 = row_col_notation(args)
      return if [row1, col1, row2, col2].include?(nil)

      # Ignore max print area since this is the same as no print area for Excel.
      if row1 == 0 && col1 == 0 && row2 == ROW_MAX - 1 && col2 == COL_MAX - 1
        return
      end

      # Build up the print area range "=Sheet2!R1C1:R2C1"
      @page_setup.print_area = convert_name_area(row1, col1, row2, col2)
    end

    #
    # Set the worksheet zoom factor in the range <tt>10 <= scale <= 400</tt>:
    #
    #     worksheet1.zoom = 50
    #     worksheet2.zoom = 75
    #     worksheet3.zoom = 300
    #     worksheet4.zoom = 400
    #
    # The default zoom factor is 100. You cannot zoom to "Selection" because
    # it is calculated by Excel at run-time.
    #
    # Note, zoom=() does not affect the scale of the printed page.
    # For that you should use print_scale=().
    #
    def zoom=(scale)
      # Confine the scale to Excel's range
      if scale < 10 or scale > 400
        # carp "Zoom factor scale outside range: 10 <= zoom <= 400"
        @zoom = 100
      else
        @zoom = scale.to_i
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
    #     worksheet1.print_scale =  50
    #     worksheet2.print_scale =  75
    #     worksheet3.print_scale = 300
    #     worksheet4.print_scale = 400
    #
    # The default scale factor is 100. Note, print_scale=() does not
    # affect the scale of the visible page in Excel. For that you should
    # use zoom=().
    #
    # Note also that although it is valid to use both fit_to_pages() and
    # print_scale=() on the same worksheet only one of these options
    # can be active at a time. The last method call made will set
    # the active option.
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
      self::print_scale = (scale)
    end

     #
     # Set the option to print the worksheet in black and white.
     #
     def print_black_and_white
       @page_setup.black_white = true
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
    # The default value is 1.
    #
    #     worksheet.set_start_page(2)
    #
    def start_page=(page_start)
      @page_setup.page_start = page_start
    end

    def set_start_page(page_start)
      put_deprecate_message("#{self}.set_start_page")
      self::start_page = page_start
    end

    #
    # :call-seq:
    #  write(row, column [ , token [ , format ] ])
    #
    # Excel makes a distinction between data types such as strings, numbers,
    # blanks, formulas and hyperlinks. To simplify the process of writing
    # data the {#write()}[#method-i-write] method acts as a general alias for several more
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
    # The general rule is that if the data looks like a _something_ then
    # a _something_ is written. Here are some examples in both row-column
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
    # The +format+ parameter is optional. It should be a valid Format object,
    # See {"CELL FORMATTING"}[Format.html#label-CELL+FORMATTING]:
    #
    #     format = workbook.add_format
    #     format.set_bold
    #     format.set_color('red')
    #     format.set_align('center')
    #
    #     worksheet.write(4, 0, 'Hello', format)    # Formatted string
    #
    # The {#write()}[#method-i-write] method will ignore empty strings or +nil+ tokens unless a
    # format is also supplied. As such you needn't worry about special handling
    # for empty or nil in your data. See also the write_blank() method.
    #
    # One problem with the {#write()}[#method-i-write] method is that occasionally data looks like
    # a number but you don't want it treated as a number. For example, zip
    # codes or ID numbers often start with a leading zero.
    # If you want to write this data with leading zero(s), use write_string.
    #
    # The write methods return:
    #     0 for success.
    #
    def write(*args)
      # Check for a cell reference in A1 notation and substitute row and column
      row_col_args = row_col_notation(args)
      token = row_col_args[2] || ''

      # Match an array ref.
      if token.respond_to?(:to_ary)
        write_row(*args)
      elsif token.respond_to?(:coerce)  # Numeric
        write_number(*args)
      elsif token =~ /^\d+$/
        write_number(*args)
      # Match formula
      elsif token =~ /^=/
        write_formula(*args)
      # Match array formula
      elsif token =~ /^\{=.*\}$/
        write_formula(*args)
      # Match blank
      elsif token == ''
        row_col_args.delete_at(2)     # remove the empty string from the parameter list
        write_blank(*row_col_args)
      elsif @workbook.strings_to_urls
        # Match http, https or ftp URL
        if token =~ %r|\A[fh]tt?ps?://|
          write_url(*args)
        # Match mailto:
        elsif token =~ %r|\Amailto:|
          write_url(*args)
        # Match internal or external sheet link
        elsif token =~ %r!\A(?:in|ex)ternal:!
          write_url(*args)
        else
          write_string(*args)
        end
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
    # The write_row() method can be used to write a 1D or 2D array of data
    # in one go. This is useful for converting the results of a database
    # query into an Excel worksheet. You must pass a reference to the array
    # of data rather than the array itself. The {#write()}[#method-i-write] method is then
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
    # Note: For convenience the {#write()}[#method-i-write] method behaves in the same way as
    # write_row() if it is passed an array.
    # Therefore the following two method calls are equivalent:
    #
    #     worksheet.write_row('A1', array)    # Write a row of data
    #     worksheet.write(    'A1', array)    # Same thing
    #
    # As with all of the write methods the +format+ parameter is optional.
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
    #
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
    # Any +nil+ in the data will be ignored unless a format is applied to
    # the data, in which case a formatted blank cell will be written.
    # In either case the appropriate row or column value will still
    # be incremented.
    #
    # See also the write_arrays.rb program in the examples directory
    # of the distro.
    #
    def write_row(*args)
      # Check for a cell reference in A1 notation and substitute row and column
      row, col, tokens, *options = row_col_notation(args)
      raise "Not an array ref in call to write_row()$!" unless tokens.respond_to?(:to_ary)

      tokens.each do |token|
        # Check for nested arrays
        if token.respond_to?(:to_ary)
          write_col(row, col, token, *options)
        else
          write(row, col, token, *options)
        end
        col += 1
      end
    end

    #
    # :call-seq:
    #   write_col(row, col, array [ , format ] )
    #
    # Write a column of data starting from (row, col). Call write_row() if any of
    # the elements of the array are in turn array. This allows the writing
    # of 1D or 2D arrays of data in one go.
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
    # As with all of the write methods the +format+ parameter is optional.
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
    # Any +nil+ in the data will be ignored unless a format is applied to
    # the data, in which case a formatted blank cell will be written.
    # In either case the appropriate row or column value will still be
    # incremented.
    #
    # As noted above the {#write()}[#method-i-write] method can be used as a synonym for
    # write_row() and write_row() handles nested array refs as columns.
    # Therefore, the following two method calls are equivalent although
    # the more explicit call to write_col() would be preferable for
    # maintainability:
    #
    #     worksheet.write_col('A1', array     ) # Write a column of data
    #     worksheet.write(    'A1', [ array ] ) # Same thing
    #
    # See also the write_arrays.rb program in the examples directory of
    # the distro.
    #
    def write_col(*args)
      row, col, tokens, *options = row_col_notation(args)

      tokens.each do |token|
        # write() will deal with any nested arrays
        write(row, col, token, *options)
        row += 1
      end
    end

    #
    # :call-seq:
    #   write_comment(row, column, string, options = {})
    #
    # Write a comment to the specified row and column (zero indexed).
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
    # cell reference. See the note about {"Cell notation"}[#label-Cell+notation].
    #
    #     worksheet.write(        'C3', 'Hello')
    #     worksheet.write_comment('C3', 'This is a comment.')
    #
    # The write_comment() method will also handle strings in UTF-8 format.
    #
    #     worksheet.write_comment('C3', "日本")
    #     worksheet.write_comment('C4', 'Comment ça va')
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
    # comments_author=() method.
    #
    #     worksheet.comments_author = 'Ruby'
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
    #     worksheet.write_comment('C3', comment, :x_offset => 30)
    #
    # ===Option: y_offset
    #
    # This option is used to change the y offset, in pixels, of a comment
    # within a cell:
    #
    #     worksheet.write_comment('C3', comment, :x_offset => 30)
    #
    # You can apply as many of these options as you require.
    #
    # <b>Note about using options that adjust the position of the cell comment
    # such as start_cell, start_row, start_col, x_offset and y_offset</b>:
    # Excel only displays offset cell comments when they are displayed as
    # "visible". Excel does not display hidden cells as moved when you
    # mouse over them.
    #
    # <b>Note about row height and comments</b>. If you specify the height of a
    # row that contains a comment then WriteXLSX will adjust the
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
      row, col, string, options = row_col_notation(args)
      raise WriteXLSXInsufficientArgumentError if [row, col, string].include?(nil)

      # Check that row and col are valid and store max and min values
      check_dimensions(row, col)
      store_row_col_max_min_values(row, col)

      @has_vml = true

      # Process the properties of the cell comment.
      @comments.add(Package::Comment.new(@workbook, self, row, col, string, options))
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
    # See the note about {"Cell notation"}[#label-Cell+notation].
    # The +format+ parameter is optional.
    #
    # In general it is sufficient to use the {#write()}[#method-i-write] method.
    #
    # Note: some versions of Excel 2007 do not display the calculated values
    # of formulas written by WriteXLSX. Applying all available Service Packs
    # to Excel should fix this.
    #
    def write_number(*args)
      # Check for a cell reference in A1 notation and substitute row and column
      row, col, num, xf = row_col_notation(args)
      raise WriteXLSXInsufficientArgumentError if [row, col, num].include?(nil)

      # Check that row and col are valid and store max and min values
      check_dimensions(row, col)
      store_row_col_max_min_values(row, col)

      store_data_to_table(NumberCellData.new(self, row, col, num, xf))
    end

    #
    # :call-seq:
    #   write_string(row, column, string [, format ] )
    #
    # Write a string to the specified row and column (zero indexed).
    # +format+ is optional.
    #
    #     worksheet.write_string(0, 0, 'Your text here')
    #     worksheet.write_string('A2', 'or here')
    #
    # The maximum string size is 32767 characters. However the maximum
    # string segment that Excel can display in a cell is 1000.
    # All 32767 characters can be displayed in the formula bar.
    #
    # In general it is sufficient to use the {#write()}[#method-i-write] method.
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
      row, col, str, xf = row_col_notation(args)
      str &&= str.to_s
      raise WriteXLSXInsufficientArgumentError if [row, col, str].include?(nil)

      # Check that row and col are valid and store max and min values
      check_dimensions(row, col)
      store_row_col_max_min_values(row, col)

      index = shared_string_index(str[0, STR_MAX])

      store_data_to_table(StringCellData.new(self, row, col, index, xf))
    end

    #
    # :call-seq:
    #    write_rich_string(row, column, (string | format, string)+,  [,cell_format] )
    #
    # The write_rich_string() method is used to write strings with multiple formats.
    # The method receives string fragments prefixed by format objects. The final
    # format object is used as the cell format.
    #
    # For example to write the string "This is *bold* and this is _italic_"
    # you would use the following:
    #
    #     bold   = workbook.add_format(:bold   => 1)
    #     italic = workbook.add_format(:italic => 1)
    #
    #     worksheet.write_rich_string('A1',
    #         'This is ', bold, 'bold', ' and this is ', italic, 'italic')
    #
    # The basic rule is to break the string into fragments and put a +format+
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
    #
    # String fragments that don't have a format are given a default format.
    # So for example when writing the string "Some *bold* text"
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
    # "http://jmcnamara.github.com/excel-writer-xlsx/images/examples/rich_strings.jpg"
    #
    # As with write_sting() the maximum string size is 32767 characters.
    # See also the note about {"Cell notation"}[#label-Cell+notation].
    #
    def write_rich_string(*args)
      # Check for a cell reference in A1 notation and substitute row and column
      row, col, *rich_strings = row_col_notation(args)
      raise WriteXLSXInsufficientArgumentError if [row, col, rich_strings[0]].include?(nil)

      xf = cell_format_of_rich_string(rich_strings)

      # Check that row and col are valid and store max and min values
      check_dimensions(row, col)
      store_row_col_max_min_values(row, col)

      fragments, length = rich_strings_fragments(rich_strings)
      # can't allow 2 formats in a row
      return -4 unless fragments

      index = shared_string_index(xml_str_of_rich_string(fragments))

      store_data_to_table(StringCellData.new(self, row, col, index, xf))
    end

    #
    # :call-seq:
    #   write_blank(row, col, format)
    #
    # Write a blank cell to the specified row and column (zero indexed).
    # A blank cell is used to specify formatting without adding a string
    # or a number.
    #
    #     worksheet.write_blank(0, 0, format)
    #
    # This method is used to add formatting to cell which doesn't contain a
    # string or number value.
    #
    # A blank cell without a format serves no purpose. Therefore, we don't write
    # a BLANK record unless a format is specified. This is mainly an optimisation
    # for the write_row() and write_col() methods.
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
    # data without special treatment for +nil+ or empty string values.
    #
    # See the note about {"Cell notation"}[#label-Cell+notation].
    #
    def write_blank(*args)
      # Check for a cell reference in A1 notation and substitute row and column
      row, col, xf = row_col_notation(args)
      raise WriteXLSXInsufficientArgumentError if [row, col].include?(nil)

      # Don't write a blank cell unless it has a format
      return unless xf

      # Check that row and col are valid and store max and min values
      check_dimensions(row, col)
      store_row_col_max_min_values(row, col)

      store_data_to_table(BlankCellData.new(self, row, col, xf))
    end

    #
    # :call-seq:
    #   write_formula(row, column, formula [ , format [ , value ] ] )
    #
    # Write a formula or function to the cell specified by +row+ and +column+:
    #
    #     worksheet.write_formula(0, 0, '=$B$3 + B4')
    #     worksheet.write_formula(1, 0, '=SIN(PI()/4)')
    #     worksheet.write_formula(2, 0, '=SUM(B1:B5)')
    #     worksheet.write_formula('A4', '=IF(A3>1,"Yes", "No")')
    #     worksheet.write_formula('A5', '=AVERAGE(1, 2, 3, 4)')
    #     worksheet.write_formula('A6', '=DATEVALUE("1-Jan-2001")')
    #
    # Array formulas are also supported:
    #
    #     worksheet.write_formula('A7', '{=SUM(A1:B1*A2:B2)}')
    #
    # See also the write_array_formula() method.
    #
    # See the note about {"Cell notation"}[#label-Cell+notation].
    # For more information about writing Excel formulas see
    # {"FORMULAS AND FUNCTIONS IN EXCEL"}[#label-FORMULAS+AND+FUNCTIONS+IN+EXCEL]
    #
    # If required, it is also possible to specify the calculated value
    # of the formula. This is occasionally necessary when working with
    # non-Excel applications that don't calculate the value of the
    # formula. The calculated +value+ is added at the end of the argument list:
    #
    #     worksheet.write('A1', '=2+2', format, 4)
    #
    # However, this probably isn't something that will ever need to do.
    # If you do use this feature then do so with care.
    #
    def write_formula(*args)
      # Check for a cell reference in A1 notation and substitute row and column
      row, col, formula, format, value = row_col_notation(args)
      raise WriteXLSXInsufficientArgumentError if [row, col, formula].include?(nil)

      if formula =~ /^\{=.*\}$/
        write_array_formula(row, col, row, col, formula, format, value)
      else
        check_dimensions(row, col)
        store_row_col_max_min_values(row, col)
        formula = formula.sub(/^=/, '')

        store_data_to_table(FormulaCellData.new(self, row, col, formula, format, value))
      end
    end

    #
    # :call-seq:
    #   write_array_formula(row1, col1, row2, col2, formula [ , format [ , value ] ] )
    #
    # Write an array formula to a cell range. In Excel an array formula is a
    # formula that performs a calculation on a set of values. It can return
    # a single value or a range of values.
    #
    # An array formula is indicated by a pair of braces around the
    # formula: +{=SUM(A1:B1*A2:B2)}+. If the array formula returns a single
    # value then the +first_+ and +last_+ parameters should be the same:
    #
    #     worksheet.write_array_formula('A1:A1', '{=SUM(B1:C1*B2:C2)}')
    #
    # It this case however it is easier to just use the write_formula()
    # or {#write()}[#method-i-write] methods:
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
      row1, col1, row2, col2, formula, xf, value = row_col_notation(args)
      raise WriteXLSXInsufficientArgumentError if [row1, col1, row2, col2, formula].include?(nil)

      # Swap last row/col with first row/col as necessary
      row1, row2 = row2, row1 if row1 > row2
      col1, col2 = col2, col1 if col1 > col2

      # Check that row and col are valid and store max and min values
      check_dimensions(row2, col2)
      store_row_col_max_min_values(row2, col2)

      # Define array range
      if row1 == row2 && col1 == col2
        range = xl_rowcol_to_cell(row1, col1)
      else
        range ="#{xl_rowcol_to_cell(row1, col1)}:#{xl_rowcol_to_cell(row2, col2)}"
      end

      # Remove array formula braces and the leading =.
      formula = formula.sub(/^\{(.*)\}$/, '\1').sub(/^=/, '')

      store_data_to_table(FormulaArrayCellData.new(self, row1, col1, formula, xf, range, value))

      # Pad out the rest of the area with formatted zeroes.
      (row1..row2).each do |row|
        (col1..col2).each do |col|
          next if row == row1 && col == col1
          write_number(row, col, 0, xf)
        end
      end
    end

    #
    # The outline_settings() method is used to control the appearance of
    # outlines in Excel. Outlines are described in
    # {"OUTLINES AND GROUPING IN EXCEL"}["method-i-set_row-label-OUTLINES+AND+GROUPING+IN+EXCEL"].
    #
    # The +visible+ parameter is used to control whether or not outlines are
    # visible. Setting this parameter to 0 will cause all outlines on the
    # worksheet to be hidden. They can be unhidden in Excel by means of the
    # "Show Outline Symbols" command button. The default setting is 1 for
    # visible outlines.
    #
    #     worksheet.outline_settings(0)
    #
    # The +symbols_below+ parameter is used to control whether the row outline
    # symbol will appear above or below the outline level bar. The default
    # setting is 1 for symbols to appear below the outline level bar.
    #
    # The +symbols_right+ parameter is used to control whether the column
    # outline symbol will appear to the left or the right of the outline level
    # bar. The default setting is 1 for symbols to appear to the right of
    # the outline level bar.
    #
    # The +auto_style+parameter is used to control whether the automatic
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
    #   write_url(row, column, url [ , format, label, tip ] )
    #
    # Write a hyperlink to a URL in the cell specified by +row+ and +column+.
    # The hyperlink is comprised of two elements: the visible label and
    # the invisible link. The visible label is the same as the link unless
    # an alternative label is specified. The label parameter is optional.
    # The label is written using the {#write()}[#method-i-write] method. Therefore it is
    # possible to write strings, numbers or formulas as labels.
    #
    # The +format+ parameter is also optional, however, without a format
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
    #     worksheet.write_url(0, 0, 'ftp://www.ruby-lang.org/',  format)
    #     worksheet.write_url('A3', 'http://www.ruby-lang.org/', format)
    #     worksheet.write_url('A4', 'mailto:foo@bar.com', format)
    #
    # You can display an alternative string using the +label+ parameter:
    #
    #     worksheet.write_url(1, 0, 'http://www.ruby-lang.org/', format, 'Ruby')
    #
    # If you wish to have some other cell data such as a number or a formula
    # you can overwrite the cell using another call to write_*():
    #
    #     worksheet.write_url('A1', 'http://www.ruby-lang.org/')
    #
    #     # Overwrite the URL string with a formula. The cell is still a link.
    #     worksheet.write_formula('A1', '=1+1', format)
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
    # All of the these URI types are recognised by the {#write()}[#method-i-write] method, see above.
    #
    # Worksheet references are typically of the form Sheet1!A1. You can
    # also refer to a worksheet range using the standard Excel notation:
    # +Sheet1!A1:B2+.
    #
    # In external links the workbook and worksheet name must be separated
    # by the # character: +external:Workbook.xlsx#Sheet1!A1+.
    #
    # You can also link to a named range in the target worksheet. For
    # example say you have a named range called +my_name+ in the workbook
    # +c:\temp\foo.xlsx+ you could link to it as follows:
    #
    #     worksheet.write_url('A14', 'external:c:\temp\foo.xlsx#my_name')
    #
    # Excel requires that worksheet names containing spaces or non
    # alphanumeric characters are single quoted as follows +'Sales Data'!A1+.
    #
    # Note: WriteXLSX will escape the following characters in URLs as required
    # by Excel: \s " < > \ [ ] ` ^ { } unless the URL already contains +%xx+
    # style escapes. In which case it is assumed that the URL was escaped
    # correctly by the user and will by passed directly to Excel.
    #
    # See also, the note about {"Cell notation"}[#label-Cell+notation].
    #
    def write_url(*args)
      # Check for a cell reference in A1 notation and substitute row and column
      row, col, url, xf, str, tip = row_col_notation(args)
      xf, str = str, xf if str.respond_to?(:xf_index) || !xf.respond_to?(:xf_index)
      raise WriteXLSXInsufficientArgumentError if [row, col, url].include?(nil)

      # Check that row and col are valid and store max and min values
      check_dimensions(row, col)
      store_row_col_max_min_values(row, col)

      hyperlink = Hyperlink.factory(url, str, tip)
      store_hyperlink(row, col, hyperlink)

      if hyperlinks_count > 65_530
        raise "URL '#{url}' added but number of URLS is over Excel's limit of 65,530 URLS per worksheet."
      end

      # Write the hyperlink string.
      write_string(row, col, hyperlink.str, xf)
    end

    #
    # :call-seq:
    #   write_date_time (row, col, date_string [ , format ] )
    #
    # Write a datetime string in ISO8601 "yyyy-mm-ddThh:mm:ss.ss" format as a
    # number representing an Excel date. format is optional.
    #
    # The write_date_time() method can be used to write a date or time
    # to the cell specified by row and column:
    #
    #     worksheet.write_date_time('A1', '2004-05-13T23:20', date_format)
    #
    # The +date_string+ should be in the following format:
    #
    #     yyyy-mm-ddThh:mm:ss.sss
    #
    # This conforms to an ISO8601 date but it should be noted that the
    # full range of ISO8601 formats are not supported.
    #
    # The following variations on the +date_string+ parameter are permitted:
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
    # A date should always have a +format+, otherwise it will appear
    # as a number, see
    # {"DATES AND TIME IN EXCEL"}[#method-i-write_date_time-label-DATES+AND+TIME+IN+EXCEL]
    # and {"CELL FORMATTING"}[#label-CELL+FORMATTING].
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
    #
    # == DATES AND TIME IN EXCEL
    #
    # There are two important things to understand about dates and times in Excel:
    #
    # 1 A date/time in Excel is a real number plus an Excel number format.
    # 2 WriteXLSX doesn't automatically convert date/time strings in {#write()}[#method-i-write] to an Excel date/time.
    #
    # These two points are explained in more detail below along with some
    # suggestions on how to convert times and dates to the required format.
    #
    # === An Excel date/time is a number plus a format
    #
    # If you write a date string with {#write()}[#method-i-write] then all you will get is a string:
    #
    #     worksheet.write('A1', '02/03/04')   # !! Writes a string not a date. !!
    #
    # Dates and times in Excel are represented by real numbers, for example
    # "Jan 1 2001 12:30 AM" is represented by the number 36892.521.
    #
    # The integer part of the number stores the number of days since the epoch
    # and the fractional part stores the percentage of the day.
    #
    # A date or time in Excel is just like any other number. To have the number
    # display as a date you must apply an Excel number format to it.
    # Here are some examples.
    #
    #     #!/usr/bin/ruby -w
    #
    #     require 'write_xlsx'
    #
    #     workbook  = WriteXLSX.new('date_examples.xlsx')
    #     worksheet = workbook>add_worksheet
    #
    #     worksheet.set_column('A:A', 30)    # For extra visibility.
    #
    #     number = 39506.5
    #
    #     worksheet.write('A1', number)             #   39506.5
    #
    #     format2 = workbook.add_format(:num_format => 'dd/mm/yy')
    #     worksheet.write('A2', number, format2)    #  28/02/08
    #
    #     format3 = workbook.add_format(:num_format => 'mm/dd/yy')
    #     worksheet.write('A3', number, format3)    #  02/28/08
    #
    #     format4 = workbook.add_format(:num_format => 'd-m-yyyy')
    #     worksheet.write('A4', number, format4)    #  28-2-2008
    #
    #     format5 = workbook.add_format(:num_format => 'dd/mm/yy hh:mm')
    #     worksheet.write('A5', number, format5)    #  28/02/08 12:00
    #
    #     format6 = workbook.add_format(:num_format => 'd mmm yyyy')
    #     worksheet.write('A6', number, format6)    # 28 Feb 2008
    #
    #     format7 = workbook.add_format(:num_format => 'mmm d yyyy hh:mm AM/PM')
    #     worksheet.write('A7', number , format7)   #  Feb 28 2008 12:00 PM
    #
    # WriteXLSX doesn't automatically convert date/time strings
    #
    # WriteXLSX doesn't automatically convert input date strings into Excel's
    # formatted date numbers due to the large number of possible date formats
    # and also due to the possibility of misinterpretation.
    #
    # For example, does 02/03/04 mean March 2 2004, February 3 2004 or
    # even March 4 2002.
    #
    # Therefore, in order to handle dates you will have to convert them to
    # numbers and apply an Excel format. Some methods for converting dates are
    # listed in the next section.
    #
    # The most direct way is to convert your dates to the ISO8601
    # yyyy-mm-ddThh:mm:ss.sss date format and use the write_date_time()
    # worksheet method:
    #
    #     worksheet.write_date_time('A2', '2001-01-01T12:20', format)
    #
    # See the write_date_time() section of the documentation for more details.
    #
    # A general methodology for handling date strings with write_date_time() is:
    #
    # 1. Identify incoming date/time strings with a regex.
    # 2. Extract the component parts of the date/time using the same regex.
    # 3. Convert the date/time to the ISO8601 format.
    # 4. Write the date/time using write_date_time() and a number format.
    # For a slightly more advanced solution you can modify the {#write()}[#method-i-write] method
    # to handle date formats of your choice via the add_write_handler() method.
    # See the add_write_handler() section of the docs and the
    # write_handler3.rb and write_handler4.rb programs in the examples
    # directory of the distro.
    #
    # Converting dates and times to an Excel date or time
    #
    # The write_date_time() method above is just one way of handling dates and
    # times.
    #
    # You can also use the convert_date_time() worksheet method to convert
    # from an ISO8601 style date string to an Excel date and time number.
    #
    def write_date_time(*args)
      # Check for a cell reference in A1 notation and substitute row and column
      row, col, str, xf = row_col_notation(args)
      raise WriteXLSXInsufficientArgumentError if [row, col, str].include?(nil)

      # Check that row and col are valid and store max and min values
      check_dimensions(row, col)
      store_row_col_max_min_values(row, col)

      date_time = convert_date_time(str)

      if date_time
        store_data_to_table(NumberCellData.new(self, row, col, date_time, xf))
      else
        # If the date isn't valid then write it as a string.
        write_string(*args)
      end
    end

    #
    # :call-seq:
    #   insert_chart(row, column, chart [ , x, y, x_scale, y_scale ] )
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
    # The +x+, +y+, +x_scale+ and +y_scale+ parameters are optional.
    #
    # The parameters +x+ and +y+ can be used to specify an offset from the top
    # left hand corner of the cell specified by +row+ and +column+. The offset
    # values are in pixels.
    #
    #     worksheet1.insert_chart('E2', chart, 3, 3)
    #
    # The parameters x_scale and y_scale can be used to scale the inserted
    # image horizontally and vertically:
    #
    #     # Scale the width by 120% and the height by 150%
    #     worksheet.insert_chart('E2', chart, 0, 0, 1.2, 1.5)
    #
    def insert_chart(*args)
      # Check for a cell reference in A1 notation and substitute row and column.
      row, col, chart, x_offset, y_offset, x_scale, y_scale = row_col_notation(args)
      raise WriteXLSXInsufficientArgumentError if [row, col, chart].include?(nil)

      x_offset ||= 0
      y_offset ||= 0
      x_scale  ||= 1
      y_scale  ||= 1

      raise "Not a Chart object in insert_chart()" unless chart.is_a?(Chart) || chart.is_a?(Chartsheet)
      raise "Not a embedded style Chart object in insert_chart()" if chart.respond_to?(:embedded) && chart.embedded == 0

      if chart.already_inserted? || (chart.combined && chart.combined.already_inserted?)
        raise "Chart cannot be inserted in a worksheet more than once"
      else
        chart.already_inserted          = true
        chart.combined.already_inserted = true if chart.combined
      end

      # Use the values set with chart.set_size, if any.
      x_scale  = chart.x_scale  if chart.x_scale  != 1
      y_scale  = chart.y_scale  if chart.y_scale  != 1
      x_offset = chart.x_offset if ptrue?(chart.x_offset)
      y_offset = chart.y_offset if ptrue?(chart.y_offset)

      @charts << [row, col, chart, x_offset, y_offset, x_scale, y_scale]
    end

    #
    # :call-seq:
    #   insert_image(row, column, filename, x=0, y=0, x_scale=1, y_scale=1)
    #
    # Partially supported. Currently only works for 96 dpi images.
    #
    # This method can be used to insert a image into a worksheet. The image
    # can be in PNG, JPEG or BMP format. The x, y, x_scale and y_scale
    # parameters are optional.
    #
    #     worksheet1.insert_image('A1', 'ruby.bmp')
    #     worksheet2.insert_image('A1', '../images/ruby.bmp')
    #     worksheet3.insert_image('A1', '.c:\images\ruby.bmp')
    #
    # The parameters +x+ and +y+ can be used to specify an offset from the top
    # left hand corner of the cell specified by +row+ and +column+. The offset
    # values are in pixels.
    #
    #     worksheet1.insert_image('A1', 'ruby.bmp', 32, 10)
    #
    # The offsets can be greater than the width or height of the underlying
    # cell. This can be occasionally useful if you wish to align two or more
    # images relative to the same cell.
    #
    # The parameters +x_scale+ and +y_scale+ can be used to scale the inserted
    # image horizontally and vertically:
    #
    #     # Scale the inserted image: width x 2.0, height x 0.8
    #     worksheet.insert_image('A1', 'perl.bmp', 0, 0, 2, 0.8)
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
    #
    def insert_image(*args)
      # Check for a cell reference in A1 notation and substitute row and column.
      row, col, image, x_offset, y_offset, x_scale, y_scale = row_col_notation(args)
      raise WriteXLSXInsufficientArgumentError if [row, col, image].include?(nil)

      x_offset ||= 0
      y_offset ||= 0
      x_scale  ||= 1
      y_scale  ||= 1

      @images << [row, col, image, x_offset, y_offset, x_scale, y_scale]
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
      # Check for a cell reference in A1 notation and substitute row and column.
      row, col, formula, format, *pairs = row_col_notation(args)
      raise WriteXLSXInsufficientArgumentError if [row, col].include?(nil)

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
    # :call-seq:
    #   set_row(row [ , height, format, hidden, level, collapsed ] )
    #
    # This method can be used to change the default properties of a row.
    # All parameters apart from +row+ are optional.
    #
    # The most common use for this method is to change the height of a row:
    #
    #     worksheet.set_row(0, 20)    # Row 1 height set to 20
    #
    # If you wish to set the format without changing the height you can
    # pass +nil+ as the height parameter:
    #
    #     worksheet.set_row(0, nil, format)
    #
    # The +format+ parameter will be applied to any cells in the row that
    # don't have a format. For example
    #
    #     worksheet.set_row(0, nil, format1)      # Set the format for row 1
    #     worksheet.write('A1', 'Hello')          # Defaults to format1
    #     worksheet.write('B1', 'Hello', format2) # Keeps format2
    #
    # If you wish to define a row format in this way you should call the
    # method before any calls to {#write()}[#method-i-write]. Calling it afterwards will overwrite
    # any format that was previously specified.
    #
    # The +hidden+ parameter should be set to 1 if you wish to hide a row.
    # This can be used, for example, to hide intermediary steps in a
    # complicated calculation:
    #
    #     worksheet.set_row(0, 20,  format, 1)
    #     worksheet.set_row(1, nil, nil,    1)
    #
    # The +level+ parameter is used to set the outline level of the row.
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
    # The +hidden+ parameter can also be used to hide collapsed outlined rows
    # when used in conjunction with the +level+ parameter.
    #
    #     worksheet.set_row(1, nil, nil, 1, 1)
    #     worksheet.set_row(2, nil, nil, 1, 1)
    #
    # For collapsed outlines you should also indicate which row has the
    # collapsed + symbol using the optional +collapsed+ parameter.
    #
    #     worksheet.set_row(3, nil, nil, 0, 0, 1)
    #
    # For a more complete example see the outline.rb and outline_collapsed.rb
    # programs in the examples directory of the distro.
    #
    # Excel allows up to 7 outline levels. Therefore the +level+ parameter
    # should be in the range <tt>0 <= level <= 7</tt>.
    #
    # == OUTLINES AND GROUPING IN EXCEL
    #
    # Excel allows you to group rows or columns so that they can be hidden or
    # displayed with a single mouse click. This feature is referred to as
    # outlines.
    #
    # Outlines can reduce complex data down to a few salient sub-totals or
    # summaries.
    #
    # This feature is best viewed in Excel but the following is an ASCII
    # representation of what a worksheet with three outlines might look like.
    # Rows 3-4 and rows 7-8 are grouped at level 2. Rows 2-9 are grouped at
    # level 1. The lines at the left hand side are called outline level bars.
    #
    #             ------------------------------------------
    #      1 2 3 |   |   A   |   B   |   C   |   D   |  ...
    #             ------------------------------------------
    #       _    | 1 |   A   |       |       |       |  ...
    #      |  _  | 2 |   B   |       |       |       |  ...
    #      | |   | 3 |  (C)  |       |       |       |  ...
    #      | |   | 4 |  (D)  |       |       |       |  ...
    #      | -   | 5 |   E   |       |       |       |  ...
    #      |  _  | 6 |   F   |       |       |       |  ...
    #      | |   | 7 |  (G)  |       |       |       |  ...
    #      | |   | 8 |  (H)  |       |       |       |  ...
    #      | -   | 9 |   I   |       |       |       |  ...
    #      -     | . |  ...  |  ...  |  ...  |  ...  |  ...
    #
    # Clicking the minus sign on each of the level 2 outlines will collapse
    # and hide the data as shown in the next figure. The minus sign changes
    # to a plus sign to indicate that the data in the outline is hidden.
    #
    #             ------------------------------------------
    #      1 2 3 |   |   A   |   B   |   C   |   D   |  ...
    #             ------------------------------------------
    #       _    | 1 |   A   |       |       |       |  ...
    #      |     | 2 |   B   |       |       |       |  ...
    #      | +   | 5 |   E   |       |       |       |  ...
    #      |     | 6 |   F   |       |       |       |  ...
    #      | +   | 9 |   I   |       |       |       |  ...
    #      -     | . |  ...  |  ...  |  ...  |  ...  |  ...
    #
    # Clicking on the minus sign on the level 1 outline will collapse the
    # remaining rows as follows:
    #
    #             ------------------------------------------
    #      1 2 3 |   |   A   |   B   |   C   |   D   |  ...
    #             ------------------------------------------
    #            | 1 |   A   |       |       |       |  ...
    #      +     | . |  ...  |  ...  |  ...  |  ...  |  ...
    #
    # Grouping in WritXLSX is achieved by setting the outline level via the
    # set_row() and set_column() worksheet methods:
    #
    #     set_row(row, height, format, hidden, level, collapsed)
    #     set_column(first_col, last_col, width, format, hidden, level, collapsed)
    #
    # The following example sets an outline level of 1 for rows 1 and 2
    # (zero-indexed) and columns B to G. The parameters $height and $XF are
    # assigned default values since they are undefined:
    #
    #     worksheet.set_row(1, nil, nil, 0, 1)
    #     worksheet.set_row(2, nil, nil, 0, 1)
    #     worksheet.set_column('B:G', nil, nil, 0, 1)
    #
    # Excel allows up to 7 outline levels. Therefore the +level+ parameter
    # should be in the range <tt>0 <= $level <= 7</tt>.
    #
    # Rows and columns can be collapsed by setting the +hidden+ flag for the
    # hidden rows/columns and setting the +collapsed+ flag for the row/column
    # that has the collapsed + symbol:
    #
    #     worksheet.set_row(1, nil, nil, 1, 1)
    #     worksheet.set_row(2, nil, nil, 1, 1)
    #     worksheet.set_row(3, nil, nil, 0, 0, 1)          # Collapsed flag.
    #
    #     worksheet.set_column('B:G', nil, nil, 1, 1)
    #     worksheet.set_column('H:H', nil, nil, 0, 0, 1)   # Collapsed flag.
    #
    # Note: Setting the $collapsed flag is particularly important for
    # compatibility with OpenOffice.org and Gnumeric.
    #
    # For a more complete example see the outline.rb and outline_collapsed.rb
    # programs in the examples directory of the distro.
    #
    # Some additional outline properties can be set via the outline_settings()
    # worksheet method, see above.
    #
    def set_row(*args)
      return unless args[0]
      row = args[0]
      height = args[1] || @default_height
      xf     = args[2]
      hidden = args[3] || 0
      level  = args[4] || 0
      collapsed = args[5] || 0

      # Get the default row height.
      default_height = @default_row_height

      # Use min col in check_dimensions. Default to 0 if undefined.
      min_col = @dim_colmin || 0

      # Check that row and col are valid and store max and min values.
      check_dimensions(row, min_col)
      store_row_col_max_min_values(row, min_col)

      height ||= default_height

      # If the height is 0 the row is hidden and the height is the default.
      if height == 0
        hidden = 1
        height = default_height
      end

      # Set the limits for the outline levels (0 <= x <= 7).
      level = 0 if level < 0
      level = 7 if level > 7

      @outline_row_level = level if level > @outline_row_level

      # Store the row properties.
      @set_rows[row] = [height, xf, hidden, level, collapsed]

      # Store the row change to allow optimisations.
      @row_size_changed = true

      height = 0 if ptrue?(hidden)

      # Store the row sizes for use when calculating image vertices.
      @row_sizes[row] = height
    end

    #
    # Set the default row properties
    #
    def set_default_row(height = nil, zero_height = nil)
      height      ||= @original_row_height
      zero_height ||= 0

      if height != @original_row_height
        @default_row_height = height

        # Store the row change to allow optimisations.
        @row_size_changed = 1
      end

      if ptrue?(zero_height)
        @default_row_zeroed = 1
      end
    end

    #
    # merge_range(first_row, first_col, last_row, last_col, string, format)
    #
    # Merge a range of cells. The first cell should contain the data and the
    # others should be blank. All cells should contain the same format.
    #
    # The merge_range() method allows you to merge cells that contain other
    # types of alignment in addition to the merging:
    #
    #     format = workbook.add_format(
    #         :border => 6,
    #         :valign => 'vcenter',
    #         :align  => 'center'
    #     )
    #
    #     worksheet.merge_range('B3:D4', 'Vertical and horizontal', format)
    #
    # merge_range() writes its +token+ argument using the worksheet
    # {#write()}[#method-i-write] method. Therefore it will handle numbers,
    # strings, formulas or urls as required. If you need to specify the
    # required write_*() method use the merge_range_type() method, see below.
    #
    # The full possibilities of this method are shown in the merge3.rb to
    # merge6.rb programs in the examples directory of the distribution.
    #
    def merge_range(*args)
      row_first, col_first, row_last, col_last, string, format, *extra_args = row_col_notation(args)

      raise "Incorrect number of arguments" if [row_first, col_first, row_last, col_last, format].include?(nil)
      raise "Fifth parameter must be a format object" unless format.respond_to?(:xf_index)
      raise "Can't merge single cell" if row_first == row_last && col_first == col_last

      # Swap last row/col with first row/col as necessary
      row_first,  row_last  = row_last,  row_first  if row_first > row_last
      col_first, col_last = col_last, col_first if col_first > col_last

      # Check that column number is valid and store the max value
      check_dimensions(row_last, col_last)
      store_row_col_max_min_values(row_last, col_last)

      # Store the merge range.
      @merge << [row_first, col_first, row_last, col_last]

      # Write the first cell
      write(row_first, col_first, string, format, *extra_args)

      # Pad out the rest of the area with formatted blank cells.
      write_formatted_blank_to_area(row_first, row_last, col_first, col_last, format)
    end

    #
    # Same as merge_range() above except the type of
    # {#write()}[#method-i-write] is specified.
    #
    # The merge_range() method, see above, uses write() to insert the required
    # data into to a merged range. However, there may be times where this
    # isn't what you require so as an alternative the merge_range_type ()
    # method allows you to specify the type of data you wish to write.
    # For example:
    #
    #     worksheet.merge_range_type('number',  'B2:C2', 123,    format1)
    #     worksheet.merge_range_type('string',  'B4:C4', 'foo',  format2)
    #     worksheet.merge_range_type('formula', 'B6:C6', '=1+2', format3)
    #
    # The +type+ must be one of the following, which corresponds to a write_*()
    # method:
    #
    #     'number'
    #     'string'
    #     'formula'
    #     'array_formula'
    #     'blank'
    #     'rich_string'
    #     'date_time'
    #     'url'
    #
    # Any arguments after the range should be whatever the appropriate method
    # accepts:
    #
    #     worksheet.merge_range_type('rich_string', 'B8:C8',
    #                                   'This is ', bold, 'bold', format4)
    #
    # Note, you must always pass a format object as an argument, even if it is
    # a default format.
    #
    def merge_range_type(type, *args)
      case type
      when 'array_formula', 'blank', 'rich_string'
        row_first, col_first, row_last, col_last, *others = row_col_notation(args)
        format = others.pop
      else
        row_first, col_first, row_last, col_last, token, format, *others = row_col_notation(args)
      end

      raise "Format object missing or in an incorrect position" unless format.respond_to?(:xf_index)
      raise "Can't merge single cell" if row_first == row_last && col_first == col_last

      # Swap last row/col with first row/col as necessary
      row_first, row_last = row_last, row_first if row_first > row_last
      col_first, col_last = col_last, col_first if col_first > col_last

      # Check that column number is valid and store the max value
      check_dimensions(row_last, col_last)
      store_row_col_max_min_values(row_last, col_last)

      # Store the merge range.
      @merge << [row_first, col_first, row_last, col_last]

      # Write the first cell
      case type
      when 'blank', 'rich_string', 'array_formula'
        others << format
      end

      case type
      when 'string'
        write_string(row_first, col_first, token, format, *others)
      when 'number'
        write_number(row_first, col_first, token, format, *others)
      when 'blank'
        write_blank(row_first, col_first, *others)
      when 'date_time'
        write_date_time(row_first, col_first, token, format, *others)
      when 'rich_string'
        write_rich_string(row_first, col_first, *others)
      when 'url'
        write_url(row_first, col_first, token, format, *others)
      when 'formula'
        write_formula(row_first, col_first, token, format, *others)
      when 'array_formula'
        write_formula_array(row_first, col_first, *others)
      else
        raise "Unknown type '#{type}'"
      end

      # Pad out the rest of the area with formatted blank cells.
      write_formatted_blank_to_area(row_first, row_last, col_first, col_last, format)
    end

    #
    # :call-seq:
    #   conditional_formatting(cell_or_cell_range, options)
    #
    # Conditional formatting is a feature of Excel which allows you to apply a
    # format to a cell or a range of cells based on a certain criteria.
    #
    # For example the following criteria is used to highlight cells >= 50 in
    # red in the conditional_format.rb example from the distro.
    #
    #     worksheet.conditional_formatting('A1:J10',
    #         {
    #             :type     => 'cell',
    #             :criteria => '>=',
    #             :value    => 50,
    #             :format   => format1
    #         }
    #     )
    #
    # http://jmcnamara.github.com/excel-writer-xlsx/images/examples/conditional_example.jpg
    #
    # The conditional_formatting method is used to apply formatting based
    # on user defined criteria to an write_xlsx file.
    #
    # It can be applied to a single cell or a range of cells.
    # You can pass 3 parameters such as (+row+, +col+, {...})
    # or 5 parameters such as (+first_row+, +first_col+, +last_row+, +last_col+, {...}).
    # You can also use A1 style notation. For example:
    #
    #     worksheet.conditional_formatting( 0, 0,       {...} )
    #     worksheet.conditional_formatting( 0, 0, 4, 1, {...} )
    #
    #     # Which are the same as:
    #
    #     worksheet.conditional_formatting( 'A1',       {...} )
    #     worksheet.conditional_formatting( 'A1:B5',    {...} )
    #
    #
    # Using A1 style notation is is also possible to specify
    # non-contiguous ranges, separated by a comma. For example:
    #
    #     worksheet.conditional_formatting( 'A1:D5,A8:D12', {...} )
    # The last parameter in conditional_formatting must be a hash containing
    # the parameters that describe the type and style of the data validation.
    #
    # The main parameters are:
    #
    #     :type
    #     :format
    #     :criteria
    #     :value
    #     :minimum
    #     :maximum
    #
    # Other, less commonly used parameters are:
    #
    #     :min_type
    #     :mid_type
    #     :max_type
    #     :min_value
    #     :mid_value
    #     :max_value
    #     :min_color
    #     :mid_color
    #     :max_color
    #     :bar_color
    #
    # Additional parameters which are used for specific conditional format types
    # are shown in the relevant sections below.
    #
    # === :type
    #
    # This parameter is passed in a hash to conditional_formatting.
    #
    # The +:type+ parameter is used to set the type of conditional formatting
    # that you wish to apply. It is always required and it has no default value.
    # Allowable type values and their associated parameters are:
    #
    #      Type             Parameters
    #     ======            ==========
    #     'cell'            :criteria
    #                       :value
    #                       :minimum
    #                       :maximum
    #
    #     'date'            :criteria
    #                       :value
    #                       :minimum
    #                       :maximum
    #
    #     'time_period'     :criteria
    #
    #     'text'            :criteria
    #                       :value
    #
    #     'average'         :criteria
    #
    #     'duplicate'       (none)
    #
    #     'unique'          (none)
    #
    #     'top'             :criteria
    #                       :value
    #
    #     'bottom'          :criteria
    #                       :value
    #
    #     'blanks'          (none)
    #
    #     'no_blanks'       (none)
    #
    #     'errors'          (none)
    #
    #     'no_errors'       (none)
    #
    #     '2_color_scale'   (none)
    #
    #     '3_color_scale'   (none)
    #
    #     'data_bar'        (none)
    #
    #     'formula'         :criteria
    #
    # All conditional formatting types have a format parameter, see below.
    # Other types and parameters such as icon sets will be added in time.
    #
    # === :type => 'cell'
    #
    # This is the most common conditional formatting type. It is used when
    # a format is applied to a cell based on a simple criterion. For example:
    #
    #     worksheet.conditional_formatting( 'A1',
    #         {
    #             :type     => 'cell',
    #             :criteria => 'greater than',
    #             :value    => 5,
    #             :format   => red_format
    #         }
    #     )
    # Or, using the between criteria:
    #
    #     worksheet.conditional_formatting( 'C1:C4',
    #         {
    #             :type     => 'cell',
    #             :criteria => 'between',
    #             :minimum  => 20,
    #             :maximum  => 30,
    #             :format   => green_format
    #         }
    #     )
    # === :criteria
    #
    # The +:criteria+ parameter is used to set the criteria by which the cell data
    # will be evaluated. It has no default value. The most common criteria
    # as applied to { type => 'cell' } are:
    #
    #     'between'
    #     'not between'
    #     'equal to'                  |  '=='  |  '='
    #     'not equal to'              |  '!='  |  '<>'
    #     'greater than'              |  '>'
    #     'less than'                 |  '<'
    #     'greater than or equal to'  |  '>='
    #     'less than or equal to'     |  '<='
    #
    # You can either use Excel's textual description strings,
    # in the first column above, or the more common symbolic alternatives.
    #
    # Additional criteria which are specific to other conditional format types
    # are shown in the relevant sections below.
    #
    # === :value
    #
    # The +:value+ is generally used along with the criteria parameter to set the
    # rule by which the cell data will be evaluated.
    #
    #     :type     => 'cell',
    #     :criteria => '>',
    #     :value    => 5
    #     :format   => format
    #
    # The +:value+ property can also be an cell reference.
    #
    #     :type     => 'cell',
    #     :criteria => '>',
    #     :value    => '$C$1',
    #     :format   => format
    #
    # === :format
    #
    # The +:format+ parameter is used to specify the format that will be applied
    # to the cell when the conditional formatting criterion is met.
    # The format is created using the add_format method in the same way as cell
    # formats:
    #
    #     format = workbook.add_format( :bold => 1, :italic => 1 )
    #
    #     worksheet.conditional_formatting( 'A1',
    #         {
    #             :type     => 'cell',
    #             :criteria => '>',
    #             :value    => 5
    #             :format   => format
    #         }
    #     )
    #
    # The conditional format follows the same rules as in Excel:
    # it is superimposed over the existing cell format and not all font and
    # border properties can be modified. Font properties that can't be modified
    # are font name, font size, superscript and subscript.
    # The border property that cannot be modified is diagonal borders.
    #
    # Excel specifies some default formats to be used with conditional
    # formatting. You can replicate them using the following write_xlsx formats:
    #
    #     # Light red fill with dark red text.
    #
    #     format1 = workbook.add_format(
    #       :bg_color => '#FFC7CE',
    #       :color    => '#9C0006'
    #     )
    #
    #     # Light yellow fill with dark yellow text.
    #
    #     format2 = workbook.add_format(
    #       :bg_color => '#FFEB9C',
    #       :color    => '#9C6500'
    #     )
    #
    #     # Green fill with dark green text.
    #
    #     format3 = workbook.add_format(
    #       :bg_color => '#C6EFCE',
    #       :color    => '#006100'
    #     )
    #
    # === :minimum
    #
    # The +:minimum+ parameter is used to set the lower limiting value when the
    # +:criteria+ is either 'between' or 'not between':
    #
    #     :validate => 'integer',
    #     :criteria => 'between',
    #     :minimum  => 1,
    #     :maximum  => 100
    #
    # === :maximum
    #
    # The +:maximum+ parameter is used to set the upper limiting value when the
    # +:criteria+ is either 'between' or 'not between'. See the previous example.
    #
    # === :type => 'date'
    #
    # The date type is the same as the cell type and uses the same criteria
    # and values. However it allows the value, minimum and maximum properties
    # to be specified in the ISO8601 yyyy-mm-ddThh:mm:ss.sss date format which
    # is detailed in the write_date_time() method.
    #
    #     worksheet.conditional_formatting( 'A1:A4',
    #         {
    #             :type     => 'date',
    #             :criteria => 'greater than',
    #             :value    => '2011-01-01T',
    #             :format   => format
    #         }
    #     )
    #
    # === :type => 'time_period'
    #
    # The time_period type is used to specify Excel's "Dates Occurring" style
    # conditional format.
    #
    #     worksheet.conditional_formatting( 'A1:A4',
    #         {
    #             :type     => 'time_period',
    #             :criteria => 'yesterday',
    #             :format   => format
    #         }
    #     )
    #
    # The period is set in the criteria and can have one of the following
    # values:
    #
    #         :criteria => 'yesterday',
    #         :criteria => 'today',
    #         :criteria => 'last 7 days',
    #         :criteria => 'last week',
    #         :criteria => 'this week',
    #         :criteria => 'next week',
    #         :criteria => 'last month',
    #         :criteria => 'this month',
    #         :criteria => 'next month'
    #
    # === :type => 'text'
    #
    # The text type is used to specify Excel's "Specific Text" style conditional
    # format. It is used to do simple string matching using the criteria and
    # value parameters:
    #
    #     worksheet.conditional_formatting( 'A1:A4',
    #         {
    #             :type     => 'text',
    #             :criteria => 'containing',
    #             :value    => 'foo',
    #             :format   => format
    #         }
    #     )
    #
    # The criteria can have one of the following values:
    #
    #     :criteria => 'containing',
    #     :criteria => 'not containing',
    #     :criteria => 'begins with',
    #     :criteria => 'ends with'
    #
    # The value parameter should be a string or single character.
    #
    # === :type => 'average'
    #
    # The average type is used to specify Excel's "Average" style conditional
    # format.
    #
    #     worksheet.conditional_formatting( 'A1:A4',
    #         {
    #             :type     => 'average',
    #             :criteria => 'above',
    #             :format   => format
    #         }
    #     )
    #
    # The type of average for the conditional format range is specified by the
    # criteria:
    #
    #     :criteria => 'above',
    #     :criteria => 'below',
    #     :criteria => 'equal or above',
    #     :criteria => 'equal or below',
    #     :criteria => '1 std dev above',
    #     :criteria => '1 std dev below',
    #     :criteria => '2 std dev above',
    #     :criteria => '2 std dev below',
    #     :criteria => '3 std dev above',
    #     :criteria => '3 std dev below'
    #
    # === :type => 'duplicate'
    #
    # The duplicate type is used to highlight duplicate cells in a range:
    #
    #     worksheet.conditional_formatting( 'A1:A4',
    #         {
    #             :type     => 'duplicate',
    #             :format   => format
    #         }
    #     )
    #
    # === :type => 'unique'
    #
    # The unique type is used to highlight unique cells in a range:
    #
    #     worksheet.conditional_formatting( 'A1:A4',
    #         {
    #             :type     => 'unique',
    #             :format   => format
    #         }
    #     )
    #
    # === :type => 'top'
    #
    # The top type is used to specify the top n values by number or percentage
    # in a range:
    #
    #     worksheet.conditional_formatting( 'A1:A4',
    #         {
    #             :type     => 'top',
    #             :value    => 10,
    #             :format   => format
    #         }
    #     )
    #
    # The criteria can be used to indicate that a percentage condition is
    # required:
    #
    #     worksheet.conditional_formatting( 'A1:A4',
    #         {
    #             :type     => 'top',
    #             :value    => 10,
    #             :criteria => '%',
    #             :format   => format
    #         }
    #     )
    #
    # === :type => 'bottom'
    #
    # The bottom type is used to specify the bottom n values by number or
    # percentage in a range.
    #
    # It takes the same parameters as top, see above.
    #
    # === :type => 'blanks'
    #
    # The blanks type is used to highlight blank cells in a range:
    #
    #     worksheet.conditional_formatting( 'A1:A4',
    #         {
    #             :type     => 'blanks',
    #             :format   => format
    #         }
    #     )
    #
    # === :type => 'no_blanks'
    #
    # The no_blanks type is used to highlight non blank cells in a range:
    #
    #     worksheet.conditional_formatting( 'A1:A4',
    #         {
    #             :type     => 'no_blanks',
    #             :format   => format
    #         }
    #     )
    #
    # === :type => 'errors'
    #
    # The errors type is used to highlight error cells in a range:
    #
    #     worksheet.conditional_formatting( 'A1:A4',
    #         {
    #             :type     => 'errors',
    #             :format   => format
    #         }
    #     )
    #
    # === :type => 'no_errors'
    #
    # The no_errors type is used to highlight non error cells in a range:
    #
    #     worksheet.conditional_formatting( 'A1:A4',
    #         {
    #             :type     => 'no_errors',
    #             :format   => format
    #         }
    #     )
    #
    # === :type => '2_color_scale'
    #
    # The 2_color_scale type is used to specify Excel's "2 Color Scale" style
    # conditional format.
    #
    #     worksheet.conditional_formatting( 'A1:A12',
    #         {
    #             :type  => '2_color_scale'
    #         }
    #     )
    #
    # At the moment only the default colors and properties can be used. These
    # will be extended in time.
    #
    # === :type => '3_color_scale'
    #
    # The 3_color_scale type is used to specify Excel's "3 Color Scale" style
    # conditional format.
    #
    #     worksheet.conditional_formatting( 'A1:A12',
    #         {
    #             :type  => '3_color_scale'
    #         }
    #     )
    #
    # At the moment only the default colors and properties can be used.
    # These will be extended in time.
    #
    # === :type => 'data_bar'
    #
    # The data_bar type is used to specify Excel's "Data Bar" style conditional
    # format.
    #
    #     worksheet.conditional_formatting( 'A1:A12',
    #         {
    #             :type  => 'data_bar',
    #         }
    #     )
    #
    # At the moment only the default colors and properties can be used. These
    # will be extended in time.
    #
    # === :type => 'formula'
    #
    # The formula type is used to specify a conditional format based on
    # a user defined formula:
    #
    #     worksheet.conditional_formatting( 'A1:A4',
    #         {
    #             :type     => 'formula',
    #             :criteria => '=$A$1 > 5',
    #             :format   => format
    #         }
    #     )
    #
    # The formula is specified in the criteria.
    #
    # === :min_type, :mid_type, :max_type
    #
    # The min_type and max_type properties are available when the conditional
    # formatting type is 2_color_scale, 3_color_scale or data_bar. The mid_type
    # is available for 3_color_scale. The properties are used as follows:
    #
    #     worksheet.conditional_formatting( 'A1:A12',
    #         {
    #             :type      => '2_color_scale',
    #             :min_type  => 'percent',
    #             :max_type  => 'percent'
    #         }
    #     )
    #
    # The available min/mid/max types are:
    #
    #     'num'
    #     'percent'
    #     'percentile'
    #     'formula'
    #
    # === :min_value, :mid_value, :max_value
    #
    # The +:min_value+ and +:max_value+ properties are available when the
    # conditional formatting type is 2_color_scale, 3_color_scale or
    # data_bar. The +:mid_value+ is available for 3_color_scale. The properties
    # are used as follows:
    #
    #     worksheet.conditional_formatting( 'A1:A12',
    #         {
    #             :type       => '2_color_scale',
    #             :min_value  => 10,
    #             :max_value  => 90
    #         }
    #     )
    #
    # === :min_color, :mid_color, :max_color, :bar_color
    #
    # The min_color and max_color properties are available when the conditional
    # formatting type is 2_color_scale, 3_color_scale or data_bar. The mid_color
    # is available for 3_color_scale. The properties are used as follows:
    #
    #     worksheet.conditional_formatting( 'A1:A12',
    #         {
    #             ;type      => '2_color_scale',
    #             :min_color => "#C5D9F1",
    #             :max_color => "#538ED5"
    #         }
    #     )
    #
    # The color can be specifies as an Excel::Writer::XLSX color index or,
    # more usefully, as a HTML style RGB hex number, as shown above.
    #
    # === Conditional Formatting Examples
    #
    # === Example 1. Highlight cells greater than an integer value.
    #
    #     worksheet.conditional_formatting( 'A1:F10',
    #         {
    #             :type     => 'cell',
    #             :criteria => 'greater than',
    #             :value    => 5,
    #             :format   => format
    #         }
    #     )
    # === Example 2. Highlight cells greater than a value in a reference cell.
    #
    #     worksheet.conditional_formatting( 'A1:F10',
    #         {
    #             :type     => 'cell',
    #             :criteria => 'greater than',
    #             :value    => '$H$1',
    #             :format   => format
    #         }
    #     )
    # === Example 3. Highlight cells greater than a certain date:
    #
    #     worksheet.conditional_formatting( 'A1:F10',
    #         {
    #             :type     => 'date',
    #             :criteria => 'greater than',
    #             :value    => '2011-01-01T',
    #             :format   => format
    #         }
    #     )
    # === Example 4. Highlight cells with a date in the last seven days:
    #
    #     worksheet.conditional_formatting( 'A1:F10',
    #         {
    #             :type     => 'time_period',
    #             :criteria => 'last 7 days',
    #             :format   => format
    #         }
    #     )
    # === Example 5. Highlight cells with strings starting with the letter b:
    #
    #     worksheet.conditional_formatting( 'A1:F10',
    #         {
    #             :type     => 'text',
    #             :criteria => 'begins with',
    #             :value    => 'b',
    #             :format   => format
    #         }
    #     )
    # === Example 6. Highlight cells that are 1 std deviation above the average for the range:
    #
    #     worksheet.conditional_formatting( 'A1:F10',
    #         {
    #             :type     => 'average',
    #             :format   => format
    #         }
    #     )
    # === Example 7. Highlight duplicate cells in a range:
    #
    #     worksheet.conditional_formatting( 'A1:F10',
    #         {
    #             :type     => 'duplicate',
    #             :format   => format
    #         }
    #     )
    # === Example 8. Highlight unique cells in a range.
    #
    #     worksheet.conditional_formatting( 'A1:F10',
    #         {
    #             :type     => 'unique',
    #             :format   => format
    #         }
    #     )
    # === Example 9. Highlight the top 10 cells.
    #
    #     worksheet.conditional_formatting( 'A1:F10',
    #         {
    #             :type     => 'top',
    #             :value    => 10,
    #             :format   => format
    #         }
    #     )
    # === Example 10. Highlight blank cells.
    #
    #     worksheet.conditional_formatting( 'A1:F10',
    #         {
    #             :type     => 'blanks',
    #             :format   => format
    #         }
    #     )
    # See also the conditional_format.rb example program in EXAMPLES.
    #
    def conditional_formatting(*args)
      cond_format = Package::ConditionalFormat.factory(self, *args)
      @cond_formats[cond_format.range] ||= []
      @cond_formats[cond_format.range] << cond_format
    end

    #
    # :call-seq:
    #    add_table(row1, col1, row2, col2, properties)
    #
    # Add an Excel table to a worksheet.
    #
    # The add_table() method is used to group a range of cells into
    # an Excel Table.
    #
    #   worksheet.add_table('B3:F7', { ... } )
    #
    # This method contains a lot of parameters and is described
    # in detail in a section
    # {"TABLES IN EXCEL"}[#method-i-add_table-label-TABLES+IN+EXCEL].
    #
    # See also the tables.rb program in the examples directory of the distro
    #
    # ==TABLES IN EXCEL
    #
    # Tables in Excel are a way of grouping a range of cells into a single
    # entity that has common formatting or that can be referenced from
    # formulas. Tables can have column headers, autofilters, total rows,
    # column formulas and default formatting.
    #
    # http://jmcnamara.github.com/excel-writer-xlsx/images/examples/tables.jpg
    #
    # For more information see "An Overview of Excel Tables"
    # http://office.microsoft.com/en-us/excel-help/overview-of-excel-tables-HA010048546.aspx.
    #
    # Tables are added to a worksheet using the add_table() method:
    #
    #     worksheet.add_table('B3:F7', parameters)
    #
    # The data range can be specified in 'A1' or 'row/col' notation (see also
    # the note about
    # {"Cell notation"}[#label-Cell+notation] for more information.
    #
    #     worksheet.add_table('B3:F7')
    #
    #     # Same as:
    #     worksheet.add_table(2, 1, 6, 5)
    #
    # The last parameter in add_table() should be a hash ref containing the
    # parameters that describe the table options and data. The available
    # parameters are:
    #
    #         :data
    #         :autofilter
    #         :header_row
    #         :banded_columns
    #         :banded_rows
    #         :first_column
    #         :last_column
    #         :style
    #         :total_row
    #         :columns
    #         :name
    #
    # The table parameters are detailed below. There are no required parameters
    # and the hash ref isn't required if no options are specified.
    #
    # ===:data
    #
    # The +:data+ parameter can be used to specify the data in the cells of the
    # table.
    #
    #     data = [
    #         [ 'Apples',  10000, 5000, 8000, 6000 ],
    #         [ 'Pears',   2000,  3000, 4000, 5000 ],
    #         [ 'Bananas', 6000,  6000, 6500, 6000 ],
    #         [ 'Oranges', 500,   300,  200,  700 ]
    #     ]
    #
    #     worksheet.add_table('B3:F7', :data => data)
    #
    # Table data can also be written separately, as an array or individual
    # cells.
    #
    #     # These two statements are the same as the single statement above.
    #     worksheet.add_table('B3:F7')
    #     worksheet.write_col('B4', data)
    #
    # Writing the cell data separately is occasionally required when you need
    # to control the write_*() method used to populate the cells or if you
    # wish to tweak the cell formatting.
    #
    # The data structure should be an array ref of array refs holding row data
    # as shown above.
    #
    # ===:header_row
    #
    # The +:header_row+ parameter can be used to turn on or off the header row
    # in the table. It is on by default.
    #
    #     worksheet.add_table('B4:F7', :header_row => 0) # Turn header off.
    #
    # The header row will contain default captions such as Column 1, Column 2,
    # etc. These captions can be overridden using the +:columns+ parameter
    # below.
    #
    # ===:autofilter
    #
    # The +:autofilter+ parameter can be used to turn on or off the autofilter
    # in the header row. It is on by default.
    #
    #     worksheet.add_table('B3:F7', :autofilter => 0) # Turn autofilter off.
    #
    # The +:autofilter+ is only shown if the +:header_row+ is on. Filters
    # within the table are not supported.
    #
    # ===:banded_rows
    #
    # The +:banded_rows+ parameter can be used to used to create rows of
    # alternating colour in the table. It is on by default.
    #
    #     worksheet.add_table('B3:F7', :banded_rows => 0)
    #
    # ===:banded_columns
    #
    # The +:banded_columns+ parameter can be used to used to create columns
    # of alternating colour in the table. It is off by default.
    #
    #     worksheet.add_table('B3:F7', :banded_columns => 1)
    #
    # ===:first_column
    #
    # The +:first_column+ parameter can be used to highlight the first column
    # of the table. The type of highlighting will depend on the style of the
    # table. It may be bold text or a different colour. It is off by default.
    #
    #     worksheet.add_table('B3:F7', :first_column => 1)
    #
    # ===:last_column
    #
    # The +:last_column+ parameter can be used to highlight the last column
    # of the table. The type of highlighting will depend on the style of the
    # table. It may be bold text or a different colour. It is off by default.
    #
    #     worksheet.add_table('B3:F7', :last_column => 1)
    #
    # ===:style
    #
    # The +:style+ parameter can be used to set the style of the table.
    # Standard Excel table format names should be used (with matching
    # capitalisation):
    #
    #     worksheet11.add_table(
    #         'B3:F7',
    #         {
    #             :data      => data,
    #             :style     => 'Table Style Light 11'
    #         }
    #     )
    #
    # The default table style is 'Table Style Medium 9'.
    #
    # ===:name
    #
    # The +:name+ parameter can be used to set the name of the table.
    #
    # By default tables are named Table1, Table2, etc. If you override the
    # table name you must ensure that it doesn't clash with an existing table
    # name and that it follows Excel's requirements for table names.
    #
    #     worksheet.add_table('B3:F7', :name => 'SalesData')
    #
    # If you need to know the name of the table, for example to use it in a
    # formula, you can get it as follows:
    #
    #     table      = worksheet2.add_table('B3:F7')
    #     table_name = table.name
    #
    # ===:total_row
    #
    # The +:total_row+ parameter can be used to turn on the total row in the
    # last row of a table. It is distinguished from the other rows by a
    # different formatting and also with dropdown SUBTOTAL functions.
    #
    #     worksheet.add_table('B3:F7', :total_row => 1)
    #
    # The default total row doesn't have any captions or functions. These must
    # by specified via the +:columns+ parameter below.
    #
    # ===:columns
    #
    # The +:columns+ parameter can be used to set properties for columns
    # within the table.
    #
    # The sub-properties that can be set are:
    #
    #     :header
    #     :formula
    #     :total_string
    #     :total_function
    #     :format
    #
    # The column data must be specified as an array of hash. For example to
    # override the default 'Column n' style table headers:
    #
    #     worksheet.add_table(
    #         'B3:F7',
    #         {
    #             :data    => data,
    #             :columns => [
    #                 { :header => 'Product' },
    #                 { :header => 'Quarter 1' },
    #                 { :header => 'Quarter 2' },
    #                 { :header => 'Quarter 3' },
    #                 { :header => 'Quarter 4' }
    #             ]
    #         }
    #     )
    #
    # If you don't wish to specify properties for a specific column you pass
    # an empty hash and the defaults will be applied:
    #
    #             ...
    #             :columns => [
    #                 { :header => 'Product' },
    #                 { :header => 'Quarter 1' },
    #                 { },                        # Defaults to 'Column 3'.
    #                 { :header => 'Quarter 3' },
    #                 { :header => 'Quarter 4' }
    #             ]
    #             ...
    #
    # Column formulas can by applied using the formula column property:
    #
    #     worksheet8.add_table(
    #         'B3:G7',
    #         {
    #             :data    => data,
    #             :columns => [
    #                 { :header => 'Product' },
    #                 { :header => 'Quarter 1' },
    #                 { :header => 'Quarter 2' },
    #                 { :header => 'Quarter 3' },
    #                 { :header => 'Quarter 4' },
    #                 {
    #                     :header  => 'Year',
    #                     :formula => '=SUM(Table8[@[Quarter 1]:[Quarter 4]])'
    #                 }
    #             ]
    #         }
    #     )
    #
    # The Excel 2007 [#This Row] and Excel 2010 @ structural references are
    # supported within the formula.
    #
    # As stated above the total_row table parameter turns on the "Total" row
    # in the table but it doesn't populate it with any defaults. Total
    # captions and functions must be specified via the columns property and
    # the total_string and total_function sub properties:
    #
    #     worksheet10.add_table(
    #         'B3:F8',
    #         {
    #             :data      => data,
    #             :total_row => 1,
    #             :columns   => [
    #                 { :header => 'Product',   total_string   => 'Totals' },
    #                 { :header => 'Quarter 1', total_function => 'sum' },
    #                 { :header => 'Quarter 2', total_function => 'sum' },
    #                 { :header => 'Quarter 3', total_function => 'sum' },
    #                 { :header => 'Quarter 4', total_function => 'sum' }
    #             ]
    #         }
    #     )
    #
    # The supported totals row SUBTOTAL functions are:
    #
    #         average
    #         count_nums
    #         count
    #         max
    #         min
    #         std_dev
    #         sum
    #         var
    #
    # User defined functions or formulas aren't supported.
    #
    # Format can also be applied to columns:
    #
    #     currency_format = workbook.add_format(:num_format => '$#,##0')
    #
    #     worksheet.add_table(
    #         'B3:D8',
    #         {
    #             :data      => data,
    #             :total_row => 1,
    #             :columns   => [
    #                 { :header => 'Product', :total_string => 'Totals' },
    #                 {
    #                     :header         => 'Quarter 1',
    #                     :total_function => 'sum',
    #                     :format         => $currency_format
    #                 },
    #                 {
    #                     :header         => 'Quarter 2',
    #                     :total_function => 'sum',
    #                     :format         => $currency_format
    #                 }
    #             ]
    #         }
    #     )
    #
    # Standard WriteXLSX format objects can be used. However, they should be
    # limited to numerical formats. Overriding other table formatting may
    # produce inconsistent results.
    #
    def add_table(*args)
      # Table count is a member of Workbook, global to all Worksheet.
      table = Package::Table.new(self, *args)
      @tables << table
      table
    end

    #
    # :call-seq:
    #    add_sparkline(properties)
    #
    # Add sparklines to the worksheet.
    #
    # Sparklines are a feature of Excel 2010+ which allows you to add small
    # charts to worksheet cells. These are useful for showing visual trends
    # in data in a compact format.
    #
    # In WriteXLSX Sparklines can be added to cells using the add_sparkline()
    # worksheet method:
    #
    #     worksheet.add_sparkline(
    #         {
    #             :location => 'F2',
    #             :range    => 'Sheet1!A2:E2',
    #             :type     => 'column',
    #             :style    => 12
    #         }
    #     )
    #
    # http://jmcnamara.github.com/excel-writer-xlsx/images/examples/sparklines1.jpg
    #
    # Note: Sparklines are a feature of Excel 2010+ only. You can write them
    # to an XLSX file that can be read by Excel 2007 but they won't be
    # displayed.
    #
    # The add_sparkline() worksheet method is used to add sparklines to a
    # cell or a range of cells.
    #
    # The parameters to add_sparkline() must be passed in a hash.
    # The main sparkline parameters are:
    #
    #     :location        (required)
    #     :range           (required)
    #     :type
    #     :style
    #
    #     :markers
    #     :negative_points
    #     :axis
    #     :reverse
    # Other, less commonly used parameters are:
    #
    #     :high_point
    #     :low_point
    #     :first_point
    #     :last_point
    #     :max
    #     :min
    #     :empty_cells
    #     :show_hidden
    #     :date_axis
    #     :weight
    #
    #     :series_color
    #     :negative_color
    #     :markers_color
    #     :first_color
    #     :last_color
    #     :high_color
    #     :low_color
    #
    # These parameters are explained in the sections below:
    #
    # ===:location
    #
    # This is the cell where the sparkline will be displayed:
    #
    #     :location => 'F1'
    #
    # The location should be a single cell. (For multiple cells see
    # {"Grouped Sparklines"}[#method-i-add_sparkline-label-Grouped+Sparklines]
    # below).
    #
    # To specify the location in row-column notation use the
    # xl_rowcol_to_cell() function from the Writexlsx::Utility module.
    #
    #     include Writexlsx::Utility
    #     ...
    #     location => xl_rowcol_to_cell( 0, 5 ), # F1
    #
    # ===:range
    #
    # This specifies the cell data range that the sparkline will plot:
    #
    #     worksheet.add_sparkline(
    #         {
    #             :location => 'F1',
    #             :range    => 'A1:E1'
    #         }
    #     )
    #
    # The range should be a 2D array. (For 3D arrays of cells see
    # {"Grouped Sparklines"}[#method-i-add_sparkline-label-Grouped+Sparklines]
    # below).
    #
    # If range is not on the same worksheet you can specify its location using
    # the usual Excel notation:
    #
    #             Lrange => 'Sheet1!A1:E1'
    #
    # If the worksheet contains spaces or special characters you should quote
    # the worksheet name in the same way that Excel does:
    #
    #             :range => q('Monthly Data'!A1:E1)
    #
    # To specify the location in row-column notation use the xl_range() or
    # xl_range_formula() functions from the Writexlsx::Utility module.
    #
    #     include Writexlsx::Utility
    #     ...
    #     range => xl_range( 1, 1,  0, 4 ),                   # 'A1:E1'
    #     range => xl_range_formula( 'Sheet1', 0, 0,  0, 4 ), # 'Sheet1!A2:E2'
    #
    # ===:type
    #
    # Specifies the type of sparkline. There are 3 available sparkline types:
    #
    #     :line    (default)
    #     :column
    #     :win_loss
    #
    # For example:
    #
    #     {
    #         :location => 'F1',
    #         :range    => 'A1:E1',
    #         :type     => 'column'
    #     }
    #
    # ===:style
    #
    # Excel provides 36 built-in Sparkline styles in 6 groups of 6. The style
    # parameter can be used to replicate these and should be a corresponding
    # number from 1 .. 36.
    #
    #     {
    #         :location => 'A14',
    #         :range    => 'Sheet2!A2:J2',
    #         :style    => 3
    #     }
    #
    # The style number starts in the top left of the style grid and runs left
    # to right. The default style is 1. It is possible to override colour
    # elements of the sparklines using the *_color parameters below.
    #
    # ===:markers
    #
    # Turn on the markers for line style sparklines.
    #
    #     {
    #         :location => 'A6',
    #         :range    => 'Sheet2!A1:J1',
    #         :markers  => 1
    #     }
    #
    # Markers aren't shown in Excel for column and win_loss sparklines.
    #
    # ===:negative_points
    #
    # Highlight negative values in a sparkline range. This is usually required
    # with win_loss sparklines.
    #
    #     {
    #         :location        => 'A21',
    #         :range           => 'Sheet2!A3:J3',
    #         :type            => 'win_loss',
    #         :negative_points => 1
    #     }
    #
    # ===:axis
    #
    # Display a horizontal axis in the sparkline:
    #
    #     {
    #         :location => 'A10',
    #         :range    => 'Sheet2!A1:J1',
    #         :axis     => 1
    #     }
    #
    # ===:reverse
    #
    # Plot the data from right-to-left instead of the default left-to-right:
    #
    #     {
    #         :location => 'A24',
    #         :range    => 'Sheet2!A4:J4',
    #         :type     => 'column',
    #         :reverse  => 1
    #     }
    #
    # ===:weight
    #
    # Adjust the default line weight (thickness) for line style sparklines.
    #
    #      :weight => 0.25
    #
    # The weight value should be one of the following values allowed by Excel:
    #
    #     0.25  0.5   0.75
    #     1     1.25
    #     2.25
    #     3
    #     4.25
    #     6
    #
    # ===:high_point, low_point, first_point, last_point
    #
    # Highlight points in a sparkline range.
    #
    #         :high_point  => 1,
    #         :low_point   => 1,
    #         :first_point => 1,
    #         :last_point  => 1
    #
    # ===:max, min
    #
    # Specify the maximum and minimum vertical axis values:
    #
    #         :max         => 0.5,
    #         :min         => -0.5
    #
    # As a special case you can set the maximum and minimum to be for a group
    # of sparklines rather than one:
    #
    #         max         => 'group'
    # See
    # {"Grouped Sparklines"}[#method-i-add_sparkline-label-Grouped+Sparklines]
    # below.
    #
    # ===:empty_cells
    #
    # Define how empty cells are handled in a sparkline.
    #
    #     :empty_cells => 'zero',
    #
    # The available options are:
    #
    #     gaps   : show empty cells as gaps (the default).
    #     zero   : plot empty cells as 0.
    #     connect: Connect points with a line ("line" type  sparklines only).
    #
    # ===:show_hidden
    #
    # Plot data in hidden rows and columns:
    #
    #     :show_hidden => 1
    #
    # Note, this option is off by default.
    #
    # ===:date_axis
    #
    # Specify an alternative date axis for the sparkline. This is useful if
    # the data being plotted isn't at fixed width intervals:
    #
    #     {
    #         :location  => 'F3',
    #         :range     => 'A3:E3',
    #         :date_axis => 'A4:E4'
    #     }
    #
    # The number of cells in the date range should correspond to the number
    # of cells in the data range.
    #
    # ===:series_color
    #
    # It is possible to override the colour of a sparkline style using the
    # following parameters:
    #
    #     :series_color
    #     :negative_color
    #     :markers_color
    #     :first_color
    #     :last_color
    #     :high_color
    #     :low_color
    #
    # The color should be specified as a HTML style #rrggbb hex value:
    #
    #     {
    #         :location     => 'A18',
    #         :range        => 'Sheet2!A2:J2',
    #         :type         => 'column',
    #         :series_color => '#E965E0'
    #     }
    #
    # ==Grouped Sparklines
    #
    # The add_sparkline() worksheet method can be used multiple times to write
    # as many sparklines as are required in a worksheet.
    #
    # However, it is sometimes necessary to group contiguous sparklines so that
    # changes that are applied to one are applied to all. In Excel this is
    # achieved by selecting a 3D range of cells for the data range and a
    # 2D range of cells for the location.
    #
    # In WriteXLSX, you can simulate this by passing an array of values to
    # location and range:
    #
    #     {
    #         :location => [ 'A27',          'A28',          'A29'          ],
    #         :range    => [ 'Sheet2!A5:J5', 'Sheet2!A6:J6', 'Sheet2!A7:J7' ],
    #         :markers  => 1
    #     }
    #
    # ===Sparkline examples
    #
    # See the sparklines1.rb and sparklines2.rb example programs in the
    # examples directory of the distro.
    #
    def add_sparkline(param)
      @sparklines << Sparkline.new(self, param, quote_sheetname(@name))
    end

    #
    # :call-seq:
    #   insert_button(row, col, properties)
    #
    # The insert_button() method can be used to insert an Excel form button
    # into a worksheet.
    #
    # This method is generally only useful when used in conjunction with
    # the Workbook add_vba_project() method to tie the button to a macro
    # from an embedded VBA project:
    #
    #     workbook  = WriteXLSX.new('file.xlsm')
    #     ...
    #     workbook.add_vba_project('./vbaProject.bin')
    #
    #     worksheet.insert_button('C2', { :macro => 'my_macro' } )
    #
    # The properties of the button that can be set are:
    #
    #     :macro
    #     :caption
    #     :width
    #     :height
    #     :x_scale
    #     :y_scale
    #     :x_offset
    #     :y_offset
    #
    # === Option: macro
    # This option is used to set the macro that the button will invoke when
    # the user clicks on it. The macro should be included using the
    # Workbook#add_vba_project() method shown above.
    #
    #     worksheet.insert_button('C2', { :macro => 'my_macro' } )
    #
    # The default macro is +ButtonX_Click+ where X is the button number.
    #
    # ===Option: caption
    # This option is used to set the caption on the button. The default is
    # <tt>Button X</tt> where X is the button number.
    #
    #     worksheet.insert_button('C2', { :macro => 'my_macro', :caption => 'Hello' })
    #
    # ===Option: width
    # This option is used to set the width of the button in pixels.
    #
    #     worksheet.insert_button('C2', { :macro => 'my_macro', :width => 128 })
    #
    # The default button width is 64 pixels which is the width of a default cell.
    #
    # ===Option: height
    # This option is used to set the height of the button in pixels.
    #
    #     worksheet.insert_button('C2', { :macro => 'my_macro', :height => 40 })
    #
    # The default button height is 20 pixels which is the height of a default cell.
    #
    # ===Option: x_scale
    # This option is used to set the width of the button as a factor of the
    # default width.
    #
    #     worksheet.insert_button('C2', { :macro => 'my_macro', :x_scale => 2.0 })
    #
    # ===Option: y_scale
    # This option is used to set the height of the button as a factor of the
    # default height.
    #
    #     worksheet.insert_button('C2', { :macro => 'my_macro', y_:scale => 2.0 } )
    #
    # ===Option: x_offset
    # This option is used to change the x offset, in pixels, of a button
    # within a cell:
    #
    #     worksheet.insert_button('C2', { :macro => 'my_macro', :x_offset => 2 })
    #
    # ===Option: y_offset
    # This option is used to change the y offset, in pixels, of a comment
    # within a cell.
    #
    # Note: Button is the only Excel form element that is available in
    # WriteXLSX. Form elements represent a lot of work to implement and the
    # underlying VML syntax isn't very much fun.
    #
    def insert_button(*args)
      @buttons_array << button_params(*(row_col_notation(args)))
      @has_vml = 1
    end

    #
    # :call-seq:
    #   data_validation(cell_or_cell_range, options)
    #
    # Data validation is a feature of Excel which allows you to restrict
    # the data that a users enters in a cell and to display help and
    # warning messages. It also allows you to restrict input to values
    # in a drop down list.
    #
    # A typical use case might be to restrict data in a cell to integer
    # values in a certain range, to provide a help message to indicate
    # the required value and to issue a warning if the input data doesn't
    # meet the stated criteria. In WriteXLSX we could do that as follows:
    #
    #     worksheet.data_validation('B3',
    #         {
    #             :validate        => 'integer',
    #             :criteria        => 'between',
    #             :minimum         => 1,
    #             :maximum         => 100,
    #             :input_title     => 'Input an integer:',
    #             :input_message   => 'Between 1 and 100',
    #             :error_message   => 'Sorry, try again.'
    #         })
    #
    # For more information on data validation see the following Microsoft
    # support article "Description and examples of data validation in Excel":
    # http://support.microsoft.com/kb/211485.
    #
    # The following sections describe how to use the data_validation()
    # method and its various options.
    #
    # The data_validation() method is used to construct an Excel
    # data validation.
    #
    # It can be applied to a single cell or a range of cells. You can pass
    # 3 parameters such as (row, col, {...})
    # or 5 parameters such as (first_row, first_col, last_row, last_col, {...}).
    # You can also use A1 style notation. For example:
    #
    #     worksheet.data_validation( 0, 0,       {...} )
    #     worksheet.data_validation( 0, 0, 4, 1, {...} )
    #
    #     # Which are the same as:
    #
    #     worksheet.data_validation( 'A1',       {...} )
    #     worksheet.data_validation( 'A1:B5',    {...} )
    #
    # See also the note about {"Cell notation"}[#label-Cell+notation] for more information.
    #
    # The last parameter in data_validation() must be a hash ref containing
    # the parameters that describe the type and style of the data validation.
    # The allowable parameters are:
    #
    #     :validate
    #     :criteria
    #     :value | minimum | source
    #     :maximum
    #     :ignore_blank
    #     :dropdown
    #
    #     :input_title
    #     :input_message
    #     :show_input
    #
    #     :error_title
    #     :error_message
    #     :error_type
    #     :show_error
    #
    # These parameters are explained in the following sections. Most of
    # the parameters are optional, however, you will generally require
    # the three main options validate, criteria and value.
    #
    #     worksheet.data_validation('B3',
    #         {
    #             :validate => 'integer',
    #             :criteria => '>',
    #             :value    => 100
    #         })
    #
    # ===validate
    #
    # This parameter is passed in a hash ref to data_validation().
    #
    # The validate parameter is used to set the type of data that you wish
    # to validate. It is always required and it has no default value.
    # Allowable values are:
    #
    #     :any
    #     :integer
    #     :decimal
    #     :list
    #     :date
    #     :time
    #     :length
    #     :custom
    #
    # +:any+ is used to specify that the type of data is unrestricted.
    # This is the same as not applying a data validation. It is only
    # provided for completeness and isn't used very often in the
    # context of WriteXLSX.
    #
    # +:integer+ restricts the cell to integer values. Excel refers to this
    # as 'whole number'.
    #
    #     :validate => 'integer',
    #     :criteria => '>',
    #     :value    => 100,
    #
    # +:decimal+ restricts the cell to decimal values.
    #
    #     :validate => 'decimal',
    #     :criteria => '>',
    #     :value    => 38.6,
    #
    # +:list+ restricts the cell to a set of user specified values. These
    # can be passed in an array ref or as a cell range (named ranges aren't
    # currently supported):
    #
    #     :validate => 'list',
    #     :value    => ['open', 'high', 'close'],
    #     # Or like this:
    #     :value    => 'B1:B3',
    #
    # Excel requires that range references are only to cells on the same
    # worksheet.
    #
    # +:date+ restricts the cell to date values. Dates in Excel are expressed
    # as integer values but you can also pass an ISO860 style string as used
    # in write_date_time(). See also
    # {"DATES AND TIME IN EXCEL"}[#method-i-write_date_time-label-DATES+AND+TIME+IN+EXCEL]
    # for more information about working with Excel's dates.
    #
    #     :validate => 'date',
    #     :criteria => '>',
    #     :value    => 39653, # 24 July 2008
    #     # Or like this:
    #     :value    => '2008-07-24T',
    #
    # +:time+ restricts the cell to time values. Times in Excel are expressed
    # as decimal values but you can also pass an ISO860 style string as used
    # in write_date_time(). See also
    # {"DATES AND TIME IN EXCEL"}[#method-i-write_date_time-label-DATES+AND+TIME+IN+EXCEL]
    # for more information about working with Excel's times.
    #
    #     :validate => 'time',
    #     :criteria => '>',
    #     :value    => 0.5, # Noon
    #     # Or like this:
    #     :value    => 'T12:00:00',
    #
    # +:length+ restricts the cell data based on an integer string length.
    # Excel refers to this as 'Text length'.
    #
    #     :validate => 'length',
    #     :criteria => '>',
    #     :value    => 10,
    #
    # +:custom+ restricts the cell based on an external Excel formula
    # that returns a TRUE/FALSE value.
    #
    #     :validate => 'custom',
    #     :value    => '=IF(A10>B10,TRUE,FALSE)',
    #
    # ===criteria
    #
    # This parameter is passed in a hash ref to data_validation().
    #
    # The +:criteria+ parameter is used to set the criteria by which the data
    # in the cell is validated. It is almost always required except for
    # the list and custom validate options. It has no default value.
    # Allowable values are:
    #
    #     'between'
    #     'not between'
    #     'equal to'                  |  '=='  |  '='
    #     'not equal to'              |  '!='  |  '<>'
    #     'greater than'              |  '>'
    #     'less than'                 |  '<'
    #     'greater than or equal to'  |  '>='
    #     'less than or equal to'     |  '<='
    #
    # You can either use Excel's textual description strings, in the first
    # column above, or the more common symbolic alternatives. The following
    # are equivalent:
    #
    #     :validate => 'integer',
    #     :criteria => 'greater than',
    #     :value    => 100,
    #
    #     :validate => 'integer',
    #     :criteria => '>',
    #     :value    => 100,
    #
    # The list and custom validate options don't require a criteria.
    # If you specify one it will be ignored.
    #
    #     :validate => 'list',
    #     :value    => ['open', 'high', 'close'],
    #
    #     :validate => 'custom',
    #     :value    => '=IF(A10>B10,TRUE,FALSE)',
    #
    # ===:value | :minimum | :source
    #
    # This parameter is passed in a hash to data_validation().
    #
    # The value parameter is used to set the limiting value to which the
    # criteria is applied. It is always required and it has no default value.
    # You can also use the synonyms minimum or source to make the validation
    # a little clearer and closer to Excel's description of the parameter:
    #
    #     # Use 'value'
    #     :validate => 'integer',
    #     :criteria => '>',
    #     :value    => 100,
    #
    #     # Use 'minimum'
    #     :validate => 'integer',
    #     :criteria => 'between',
    #     :minimum  => 1,
    #     :maximum  => 100,
    #
    #     # Use 'source'
    #     :validate => 'list',
    #     :source   => '$B$1:$B$3',
    #
    # ===:maximum
    #
    # This parameter is passed in a hash ref to data_validation().
    #
    # The +:maximum: parameter is used to set the upper limiting value when
    # the criteria is either 'between' or 'not between':
    #
    #     :validate => 'integer',
    #     :criteria => 'between',
    #     :minimum  => 1,
    #     :maximum  => 100,
    #
    # ===:ignore_blank
    #
    # This parameter is passed in a hash ref to data_validation().
    #
    # The +:ignore_blank+ parameter is used to toggle on and off the
    # 'Ignore blank' option in the Excel data validation dialog. When the
    # option is on the data validation is not applied to blank data in the
    # cell. It is on by default.
    #
    #     :ignore_blank => 0,  # Turn the option off
    #
    # ===:dropdown
    #
    # This parameter is passed in a hash ref to data_validation().
    #
    # The +:dropdown+ parameter is used to toggle on and off the
    # 'In-cell dropdown' option in the Excel data validation dialog.
    # When the option is on a dropdown list will be shown for list validations.
    # It is on by default.
    #
    #     :dropdown => 0,      # Turn the option off
    #
    # ===:input_title
    #
    # This parameter is passed in a hash ref to data_validation().
    #
    # The +:input_title+ parameter is used to set the title of the input
    # message that is displayed when a cell is entered. It has no default
    # value and is only displayed if the input message is displayed.
    # See the input_message parameter below.
    #
    #     :input_title   => 'This is the input title',
    #
    # The maximum title length is 32 characters.
    #
    # ===:input_message
    #
    # This parameter is passed in a hash ref to data_validation().
    #
    # The +:input_message+ parameter is used to set the input message that
    # is displayed when a cell is entered. It has no default value.
    #
    #     :validate      => 'integer',
    #     :criteria      => 'between',
    #     :minimum       => 1,
    #     :maximum       => 100,
    #     :input_title   => 'Enter the applied discount:',
    #     :input_message => 'between 1 and 100',
    #
    # The message can be split over several lines using newlines, "\n" in
    # double quoted strings.
    #
    #     input_message => "This is\na test.",
    #
    # The maximum message length is 255 characters.
    #
    # ===:show_input
    #
    # This parameter is passed in a hash ref to data_validation().
    #
    # The +:show_input+ parameter is used to toggle on and off the 'Show input
    # message when cell is selected' option in the Excel data validation
    # dialog. When the option is off an input message is not displayed even
    # if it has been set using input_message. It is on by default.
    #
    #     :show_input => 0,      # Turn the option off
    #
    # ===:error_title
    #
    # This parameter is passed in a hash ref to data_validation().
    #
    # The +:error_title+ parameter is used to set the title of the error message
    # that is displayed when the data validation criteria is not met.
    # The default error title is 'Microsoft Excel'.
    #
    #     :error_title   => 'Input value is not valid',
    #
    # The maximum title length is 32 characters.
    #
    # ===:error_message
    #
    # This parameter is passed in a hash ref to data_validation().
    #
    # The +:error_message+ parameter is used to set the error message that is
    # displayed when a cell is entered. The default error message is
    # "The value you entered is not valid.\nA user has restricted values
    # that can be entered into the cell.".
    #
    #     :validate      => 'integer',
    #     :criteria      => 'between',
    #     :minimum       => 1,
    #     :maximum       => 100,
    #     :error_title   => 'Input value is not valid',
    #     :error_message => 'It should be an integer between 1 and 100',
    #
    # The message can be split over several lines using newlines, "\n"
    # in double quoted strings.
    #
    #     :input_message => "This is\na test.",
    #
    # The maximum message length is 255 characters.
    #
    # ===:error_type
    #
    # This parameter is passed in a hash ref to data_validation().
    #
    # The +:error_type+ parameter is used to specify the type of error dialog
    # that is displayed. There are 3 options:
    #
    #     'stop'
    #     'warning'
    #     'information'
    #
    # The default is 'stop'.
    #
    # ===:show_error
    #
    # This parameter is passed in a hash ref to data_validation().
    #
    # The +:show_error+ parameter is used to toggle on and off the 'Show error
    # alert after invalid data is entered' option in the Excel data validation
    # dialog. When the option is off an error message is not displayed
    # even if it has been set using error_message. It is on by default.
    #
    #     :show_error => 0,      # Turn the option off
    #
    # ===Data Validation Examples
    #
    # ===Example 1. Limiting input to an integer greater than a fixed value.
    #
    #     worksheet.data_validation('A1',
    #         {
    #             :validate        => 'integer',
    #             :criteria        => '>',
    #             :value           => 0,
    #         });
    # ===Example 2. Limiting input to an integer greater than a fixed value where the value is referenced from a cell.
    #
    #     worksheet.data_validation('A2',
    #         {
    #             :validate        => 'integer',
    #             :criteria        => '>',
    #             :value           => '=E3',
    #         });
    # ===Example 3. Limiting input to a decimal in a fixed range.
    #
    #     worksheet.data_validation('A3',
    #         {
    #             :validate        => 'decimal',
    #             :criteria        => 'between',
    #             :minimum         => 0.1,
    #             :maximum         => 0.5,
    #         });
    # ===Example 4. Limiting input to a value in a dropdown list.
    #
    #     worksheet.data_validation('A4',
    #         {
    #             :validate        => 'list',
    #             :source          => ['open', 'high', 'close'],
    #         });
    # ===Example 5. Limiting input to a value in a dropdown list where the list is specified as a cell range.
    #
    #     worksheet.data_validation('A5',
    #         {
    #             :validate        => 'list',
    #             :source          => '=$E$4:$G$4',
    #         });
    # ===Example 6. Limiting input to a date in a fixed range.
    #
    #     worksheet.data_validation('A6',
    #         {
    #             :validate        => 'date',
    #             :criteria        => 'between',
    #             :minimum         => '2008-01-01T',
    #             :maximum         => '2008-12-12T',
    #         });
    # ===Example 7. Displaying a message when the cell is selected.
    #
    #     worksheet.data_validation('A7',
    #         {
    #             :validate      => 'integer',
    #             :criteria      => 'between',
    #             :minimum       => 1,
    #             :maximum       => 100,
    #             :input_title   => 'Enter an integer:',
    #             :input_message => 'between 1 and 100',
    #         });
    # See also the data_validate.rb program in the examples directory
    # of the distro.
    #
    def data_validation(*args)
      validation = DataValidation.new(*args)
      @validations << validation unless validation.validate_none?
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
    def hide_gridlines(option = 1)
      if option == 2
        @screen_gridlines = false
      else
        @screen_gridlines = true
      end

      @page_setup.hide_gridlines(option)
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
    def print_row_col_headers(headers = true)
      @page_setup.print_row_col_headers(headers)
      # if headers
      #   @print_headers         = 1
      #   @page_setup.print_options_changed = 1
      # else
      #   @print_headers = 0
      # end
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
      @page_setup.fit_page   = true
      @page_setup.fit_width  = width
      @page_setup.fit_height  = height
      @page_setup.page_setup_changed = true
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
      row1, col1, row2, col2 = row_col_notation(args)
      return if [row1, col1, row2, col2].include?(nil)

      # Reverse max and min values if necessary.
      row1, row2 = row2, row1 if row2 < row1
      col1, col2 = col2, col1 if col2 < col1

      @autofilter_area = convert_name_area(row1, col1, row2, col2)
      @autofilter_ref  = xl_range(row1, row2, col1, col2)
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
    # Rows are hidden using the set_row() +visible+ parameter. WriteXLSX cannot
    # do this automatically since it isn't part of the file format.
    # See the autofilter.rb program in the examples directory of the distro
    # for an example.
    #
    # The conditions for the filter are specified using simple expressions:
    #
    #     worksheet.filter_column('A', 'x > 2000')
    #     worksheet.filter_column('B', 'x > 2000 and x < 5000')
    #
    # The +column+ parameter can either be a zero indexed column number or
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
    # separated by the +and+ and +or+ operators. For example:
    #
    #     'x <  2000'
    #     'x >  2000'
    #     'x == 2000'
    #     'x >  2000 and x <  5000'
    #     'x == 2000 or  x == 5000'
    #
    # Filtering of blank or non-blank data can be achieved by using a value
    # of +Blanks+ or +NonBlanks+ in the expression:
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
    # be escaped using +~+.
    #
    # The placeholder variable +x+ in the above examples can be replaced by any
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
    # Note writeExcel gem supports Top 10 style filters. These aren't
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
      breaks = args.collect do |brk|
        Array(brk)
      end.flatten
      @page_setup.hbreaks += breaks
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
      @page_setup.vbreaks += args
    end

    #
    # This method is used to make all cell comments visible when a worksheet
    # is opened.
    #
    #     worksheet.show_comments
    #
    # Individual comments can be made visible using the visible parameter of
    # the write_comment method:
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
    # This method is used to set the default author of all cell comments.
    #
    #     worksheet.comments_author = 'Ruby'
    #
    # Individual comment authors can be set using the author parameter
    # of the write_comment method.
    #
    # The default comment author is an empty string, '',
    # if no author is specified.
    #
    def comments_author=(author)
      @comments_author = author || ''
    end

    # This method is deprecated. use comments_author=().
    def set_comments_author(author)
      put_deprecate_message("#{self}.set_comments_author")
      self.comments_author = author
    end

    def has_vml?  # :nodoc:
      @has_vml
    end

    def has_header_vml?  # :nodoc:
      @has_header_vml
    end

    def has_comments? # :nodoc:
      !@comments.empty?
    end

    def has_shapes?
      @has_shapes
    end

    def is_chartsheet? # :nodoc:
      !!@is_chartsheet
    end

    def set_external_vml_links(vml_drawing_id) # :nodoc:
      @external_vml_links <<
        ['/vmlDrawing', "../drawings/vmlDrawing#{vml_drawing_id}.vml"]
    end

    def set_external_comment_links(comment_id) # :nodoc:
      @external_comment_links <<
        ['/comments',   "../comments#{comment_id}.xml"]
    end

    #
    # Set up chart/drawings.
    #
    def prepare_chart(index, chart_id, drawing_id) # :nodoc:
      drawing_type = 1

      row, col, chart, x_offset, y_offset, x_scale, y_scale  = @charts[index]
      chart.id = chart_id - 1
      x_scale ||= 0
      y_scale ||= 0

      # Use user specified dimensions, if any.
      width  = chart.width  if ptrue?(chart.width)
      height = chart.height if ptrue?(chart.height)

      width  = (0.5 + (width  * x_scale)).to_i
      height = (0.5 + (height * y_scale)).to_i

      dimensions = position_object_emus(col, row, x_offset, y_offset, width, height)

      # Set the chart name for the embedded object if it has been specified.
      name = chart.name

      # Create a Drawing object to use with worksheet unless one already exists.
      if !drawing?
        drawing = Drawing.new
        drawing.add_drawing_object(drawing_type, dimensions, 0, 0, name)
        drawing.embedded = 1

        @drawing = drawing

        @external_drawing_links << ['/drawing', "../drawings/drawing#{drawing_id}.xml" ]
      else
        @drawing.add_drawing_object(drawing_type, dimensions, 0, 0, name)
      end
      @drawing_links << ['/chart', "../charts/chart#{chart_id}.xml"]
    end

    #
    # Returns a range of data from the worksheet _table to be used in chart
    # cached data. Strings are returned as SST ids and decoded in the workbook.
    # Return nils for data that doesn't exist since Excel can chart series
    # with data missing.
    #
    def get_range_data(row_start, col_start, row_end, col_end) # :nodoc:
      # TODO. Check for worksheet limits.

      # Iterate through the table data.
      data = []
      (row_start .. row_end).each do |row_num|
        # Store nil if row doesn't exist.
        if !@cell_data_table[row_num]
          data << nil
          next
        end

        (col_start .. col_end).each do |col_num|
          if cell = @cell_data_table[row_num][col_num]
            data << cell.data
          else
            # Store nil if col doesn't exist.
            data << nil
          end
        end
      end

      return data
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
    #   +-----+----|    Object    |-----+
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
    #     col_start, row_start, col_end, row_end, x1, y1, x2, y2.
    #
    # We also calculate the absolute x and y position of the top left vertex of
    # the object. This is required for images.
    #
    #    x_abs, y_abs
    #
    # The width and height of the cells that the object occupies can be variable
    # and have to be taken into account.
    #
    # The values of col_start and row_start are passed in from the calling
    # function. The values of col_end and row_end are calculated by subtracting
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
    def position_object_pixels(col_start, row_start, x1, y1, width, height) #:nodoc:
      # Calculate the absolute x offset of the top-left vertex.
      if @col_size_changed
        x_abs = (0 .. col_start-1).inject(0) {|sum, col| sum += size_col(col)}
      else
        # Optimisation for when the column widths haven't changed.
        x_abs = @default_col_pixels * col_start
      end
      x_abs += x1

      # Calculate the absolute y offset of the top-left vertex.
      # Store the column change to allow optimisations.
      if @row_size_changed
        y_abs = (0 .. row_start-1).inject(0) {|sum, row| sum += size_row(row)}
      else
        # Optimisation for when the row heights haven't changed.
        y_abs = @default_row_pixels * row_start
      end
      y_abs += y1

      # Adjust start column for offsets that are greater than the col width.
      x1, col_start = adjust_column_offset(x1, col_start)

      # Adjust start row for offsets that are greater than the row height.
      y1, row_start = adjust_row_offset(y1, row_start)

      # Initialise end cell to the same as the start cell.
      col_end = col_start
      row_end = row_start

      width  += x1
      height += y1

      # Subtract the underlying cell widths to find the end cell of the object.
      width, col_end = adjust_column_offset(width, col_end)

      # Subtract the underlying cell heights to find the end cell of the object.
      height, row_end = adjust_row_offset(height, row_end)

      # The end vertices are whatever is left from the width and height.
      x2 = width
      y2 = height

      [col_start, row_start, x1, y1, col_end, row_end, x2, y2, x_abs, y_abs]
    end

    def comments_visible? # :nodoc:
      !!@comments_visible
    end

    def sorted_comments # :nodoc:
      @comments.sorted_comments
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
      @writer.data_element('f', formula,
                           [
                            ['t', 'array'],
                            ['ref', range]
                           ]
                           )
    end

    def date_1904? #:nodoc:
      @workbook.date_1904?
    end

    def excel2003_style? # :nodoc:
      @workbook.excel2003_style
    end

    #
    # Convert from an Excel internal colour index to a XML style #RRGGBB index
    # based on the default or user defined values in the Workbook palette.
    #
    def palette_color(index) #:nodoc:
      if index =~ /^#([0-9A-F]{6})$/i
        "FF#{$1.upcase}"
      else
        "FF#{super(index)}"
      end
    end

    def buttons_data  # :nodoc:
      @buttons_array
    end

    def header_images_data  # :nodoc:
      @header_images_array
    end

    def external_links
      [
       @external_hyper_links,
       @external_drawing_links,
       @external_vml_links,
       @external_table_links,
       @external_comment_links
      ].reject { |a| a.empty? }
    end

    def drawing_links
      [@drawing_links]
    end

    #
    # Turn the HoH that stores the comments into an array for easier handling
    # and set the external links for comments and buttons.
    #
    def prepare_vml_objects(vml_data_id, vml_shape_id, vml_drawing_id, comment_id)
      set_external_vml_links(vml_drawing_id)
      set_external_comment_links(comment_id) if has_comments?

      # The VML o:idmap data id contains a comma separated range when there is
      # more than one 1024 block of comments, like this: data="1,2".
      data = "#{vml_data_id}"
      (1 .. num_comments_block).each do |i|
        data += ",#{vml_data_id + i}"
      end
      @vml_data_id = data
      @vml_shape_id = vml_shape_id
    end

    #
    # Setup external linkage for VML header/footer images.
    #
    def prepare_header_vml_objects(vml_header_id, vml_drawing_id)
      @vml_header_id = vml_header_id
      @external_vml_links << ['/vmlDrawing', "../drawings/vmlDrawing#{vml_drawing_id}.vml"]
    end

    #
    # Set the table ids for the worksheet tables.
    #
    def prepare_tables(table_id)
      if tables_count > 0
        id = table_id
        tables.each do |table|
          table.prepare(id)

          # Store the link used for the rels file.
          @external_table_links << ['/table', "../tables/table#{id}.xml"]
          id += 1
        end
      end
      tables_count || 0
    end

    def num_comments_block
      @comments.size / 1024
    end

    def tables_count
      @tables.size
    end

    def horizontal_dpi=(val)
      @page_setup.horizontal_dpi = val
    end

    def vertical_dpi=(val)
      @page_setup.vertical_dpi = val
    end

    #
    # set the vba name for the worksheet
    #
    def set_vba_name(vba_codename = nil)
      if vba_codename
        @vba_codename = vba_codename
      else
        @vba_codename = @name
      end
    end

    private

    def hyperlinks_count
      @hyperlinks.keys.inject(0) { |s, n| s += @hyperlinks[n].keys.size }
    end

    def store_hyperlink(row, col, hyperlink)
      @hyperlinks      ||= {}
      @hyperlinks[row] ||= {}
      @hyperlinks[row][col] = hyperlink
    end

    def cell_format_of_rich_string(rich_strings)
      # If the last arg is a format we use it as the cell format.
      if rich_strings[-1].respond_to?(:xf_index)
        rich_strings.pop
      else
        nil
      end
    end

    # Convert the list of format, string tokens to pairs of (format, string)
    # except for the first string fragment which doesn't require a default
    # formatting run. Use the default for strings without a leading format.
    def rich_strings_fragments(rich_strings) # :nodoc:
      # Create a temp format with the default font for unformatted fragments.
      default = Format.new(0)

      length = 0                     # String length.
      last = 'format'
      pos  = 0

      fragments = []
      rich_strings.each do |token|
        if token.respond_to?(:xf_index)
          # Can't allow 2 formats in a row
          return nil if last == 'format' && pos > 0

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
      [fragments, length]
    end

    def xml_str_of_rich_string(fragments)
      # Create a temp XML::Writer object and use it to write the rich string
      # XML to a string.
      writer = Package::XMLWriterSimple.new

      # If the first token is a string start the <r> element.
      writer.start_tag('r') if !fragments[0].respond_to?(:xf_index)

      # Write the XML elements for the format string fragments.
      fragments.each do |token|
        if token.respond_to?(:xf_index)
          # Write the font run.
          writer.start_tag('r')
          token.write_font_rpr(writer, self)
        else
          # Write the string fragment part, with whitespace handling.
          attributes = []

          attributes << ['xml:space', 'preserve'] if token =~ /^\s/ || token =~ /\s$/
          writer.data_element('t', token, attributes)
          writer.end_tag('r')
        end
      end
      writer.string
    end

    # Pad out the rest of the area with formatted blank cells.
    def write_formatted_blank_to_area(row_first, row_last, col_first, col_last, format)
      (row_first .. row_last).each do |row|
        (col_first .. col_last).each do |col|
          next if row == row_first && col == col_first
          write_blank(row, col, format)
        end
      end
    end

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

    def adjust_column_offset(x, column)
      while x >= size_col(column)
        x -= size_col(column)
        column += 1
      end
      [x, column]
    end

    def adjust_row_offset(y, row)
      while y >= size_row(row)
        y -= size_row(row)
        row += 1
      end
      [y, row]
    end

    #
    # Calculate the vertices that define the position of a graphical object within
    # the worksheet in EMUs.
    #
    # The vertices are expressed as English Metric Units (EMUs). There are 12,700
    # EMUs per point. Therefore, 12,700 * 3 /4 = 9,525 EMUs per pixel.
    #
    def position_object_emus(col_start, row_start, x1, y1, width, height, x_dpi = 96, y_dpi = 96) #:nodoc:
      col_start, row_start, x1, y1, col_end, row_end, x2, y2, x_abs, y_abs =
        position_object_pixels(col_start, row_start, x1, y1, width, height)

      # Convert the pixel values to EMUs. See above.
      x1    = (0.5 + 9_525 * x1).to_i
      y1    = (0.5 + 9_525 * y1).to_i
      x2    = (0.5 + 9_525 * x2).to_i
      y2    = (0.5 + 9_525 * y2).to_i
      x_abs = (0.5 + 9_525 * x_abs).to_i
      y_abs = (0.5 + 9_525 * y_abs).to_i

      [col_start, row_start, x1, y1, col_end, row_end, x2, y2, x_abs, y_abs]
    end

    #
    # Convert the width of a cell from user's units to pixels. Excel rounds the
    # column width to the nearest pixel. If the width hasn't been set by the user
    # we use the default value. If the column is hidden it has a value of zero.
    #
    def size_col(col) #:nodoc:
      # Look up the cell value to see if it has been changed.
      if @col_sizes[col]
        width = @col_sizes[col]

        # Convert to pixels.
        if width == 0
          pixels = 0
        elsif width < 1
          pixels = (width * (MAX_DIGIT_WIDTH + PADDING) + 0.5).to_i
        else
          pixels = (width * MAX_DIGIT_WIDTH + 0.5).to_i + PADDING
        end
      else
        pixels = @default_col_pixels
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
        pixels = (4 / 3.0 * @default_row_height).to_i
      end
      pixels
    end

    #
    # Set up image/drawings.
    #
    def prepare_image(index, image_id, drawing_id, width, height, name, image_type, x_dpi = 96, y_dpi = 96) #:nodoc:
      x_dpi ||= 96
      y_dpi ||= 96
      drawing_type = 2
      drawing

      row, col, image, x_offset, y_offset, x_scale, y_scale = @images[index]

      width  *= x_scale
      height *= y_scale

      width  *= 96.0 / x_dpi
      height *= 96.0 / y_dpi

      dimensions = position_object_emus(col, row, x_offset, y_offset, width, height)

      # Convert from pixels to emus.
      width  = (0.5 + (width  * 9_525)).to_i
      height = (0.5 + (height * 9_525)).to_i

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
    public :prepare_image

    def prepare_header_image(image_id, width, height, name, image_type, position, x_dpi, y_dpi)
      # Strip the extension from the filename.
      body = name.dup
      body[/\.[^\.]+$/, 0] = ''

      @header_images_array << [width, height, body, position, x_dpi, y_dpi]
      @vml_drawing_links   << ['/image', "../media/image#{image_id}.#{image_type}" ]
    end
    public :prepare_header_image

    #
    # :call-seq:
    #   insert_shape(row, col, shape [ , x, y, x_scale, y_scale ] )
    #
    # Insert a shape into the worksheet.
    #
    # This method can be used to insert a Shape object into a worksheet.
    # The Shape must be created by the add_shape() Workbook method.
    #
    #   shape = workbook.add_shape(:name => 'My Shape', :type => 'plus')
    #
    #   # Configure the shape.
    #   shape.set_text('foo')
    #   ...
    #
    #   # Insert the shape into the a worksheet.
    #   worksheet.insert_shape('E2', shape)
    #
    # See add_shape() for details on how to create the Shape object
    # and Writexlsx::Shape for details on how to configure it.
    #
    # The +x+, +y+, +x_scale+ and +y_scale+ parameters are optional.
    #
    # The parameters +x+ and +y+ can be used to specify an offset
    # from the top left hand corner of the cell specified by +row+ and +col+.
    # The offset values are in pixels.
    #
    #   worksheet1.insert_shape('E2', chart, 3, 3)
    #
    # The parameters +x_scale+ and +y_scale+ can be used to scale the
    # inserted shape horizontally and vertically:
    #
    #   # Scale the width by 120% and the height by 150%
    #   worksheet.insert_shape('E2', shape, 0, 0, 1.2, 1.5)
    #
    # See also the shape*.rb programs in the examples directory of the distro.
    #
    def insert_shape(*args)
      # Check for a cell reference in A1 notation and substitute row and column.
      row_start, column_start, shape, x_offset, y_offset, x_scale, y_scale =
        row_col_notation(args)
      if [row_start, column_start, shape].include?(nil)
        raise "Insufficient arguments in insert_shape()"
      end

      shape.set_position(
                           row_start, column_start, x_offset, y_offset,
                           x_scale, y_scale
                           )
      # Assign a shape ID.
      while true
        id = shape.id || 0
        used = @shape_hash[id]

        # Test if shape ID is already used. Otherwise assign a new one.
        if !used && id != 0
          break
        else
          @last_shape_id += 1
          shape.id = @last_shape_id
        end
      end

      # Allow lookup of entry into shape array by shape ID.
      @shape_hash[shape.id] = shape.element = @shapes.size

      if ptrue?(shape.stencil)
        # Insert a copy of the shape, not a reference so that the shape is
        # used as a stencil. Previously stamped copies don't get modified
        # if the stencil is modified.
        insert = shape.dup
      else
        insert = shape
      end

      # For connectors change x/y coords based on location of connected shapes.
      insert.auto_locate_connectors(@shapes, @shape_hash)

      # Insert a link to the shape on the list of shapes. Connection to
      # the parent shape is maintained.
      @shapes << insert
      insert
    end
    public :insert_shape

    #
    # Set up drawing shapes
    #
    def prepare_shape(index, drawing_id)
      shape = @shapes[index]

      # Create a Drawing object to use with worksheet unless one already exists.
      unless drawing?
        @drawing = Drawing.new
        @drawing.embedded = 1
        @external_drawing_links << ['/drawing', "../drawings/drawing#{drawing_id}.xml"]
        @has_shapes = true
      end

      # Validate the he shape against various rules.
      shape.validate(index)
      shape.calc_position_emus(self)

      drawing_type = 3
      drawing.add_drawing_object(drawing_type, shape.dimensions, shape.name, shape)
    end
    public :prepare_shape

    #
    # This method handles the parameters passed to insert_button as well as
    # calculating the comment object position and vertices.
    #
    def button_params(row, col, params)
      button = Writexlsx::Package::Button.new

      button_number = 1 + @buttons_array.size

      # Set the button caption.
      caption = params[:caption] || "Button #{button_number}"

      button.font = { :_caption => caption }

      # Set the macro name.
      if params[:macro]
        button.macro = "[0]!#{params[:macro]}"
      else
        button.macro = "[0]!Button#{button_number}_Click"
      end

      # Ensure that a width and height have been set.
      default_width  = @default_col_pixels
      default_height = @default_row_pixels
      params[:width]  = default_width  if !params[:width]
      params[:height] = default_height if !params[:height]

      # Set the x/y offsets.
      params[:x_offset] = 0 if !params[:x_offset]
      params[:y_offset] = 0 if !params[:y_offset]

      # Scale the size of the comment box if required.
      if params[:x_scale]
        params[:width] = params[:width] * params[:x_scale]
      end
      if params[:y_scale]
        params[:height] = params[:height] * params[:y_scale]
      end

      # Round the dimensions to the nearest pixel.
      params[:width]  = (0.5 + params[:width]).to_i
      params[:height] = (0.5 + params[:height]).to_i

      params[:start_row] = row
      params[:start_col] = col

      # Calculate the positions of comment object.
      vertices = position_object_pixels(
                                        params[:start_col],
                                        params[:start_row],
                                        params[:x_offset],
                                        params[:y_offset],
                                        params[:width],
                                        params[:height]
                                        )

      # Add the width and height for VML.
      vertices << [params[:width], params[:height]]

      button.vertices = vertices

      button
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
    def write_worksheet_attributes #:nodoc:
      schema = 'http://schemas.openxmlformats.org/'
      attributes = [
                    ['xmlns',    "#{schema}spreadsheetml/2006/main"],
                    ['xmlns:r',  "#{schema}officeDocument/2006/relationships"]
                   ]
      if @excel_version == 2010
        attributes << ['xmlns:mc',     "#{schema}markup-compatibility/2006"]
        attributes << ['xmlns:x14ac',  "#{OFFICE_URL}spreadsheetml/2009/9/ac"]
        attributes << ['mc:Ignorable', 'x14ac']
      end
      attributes
    end

    #
    # Write the <sheetPr> element for Sheet level properties.
    #
    def write_sheet_pr #:nodoc:
      return unless tab_outline_fit? || vba_codename? || filter_on?

      attributes = []
      attributes << ['codeName',   @vba_codename] if vba_codename?
      attributes << ['filterMode', 1]             if filter_on?

      if tab_outline_fit?
        @writer.tag_elements('sheetPr', attributes) do
          write_tab_color
          write_outline_pr
          write_page_set_up_pr
        end
      else
        @writer.empty_tag('sheetPr', attributes)
      end
    end

    def tab_outline_fit?
      tab_color? || outline_changed? || fit_page?
    end

    #
    # Write the <pageSetUpPr> element.
    #
    def write_page_set_up_pr #:nodoc:
      @writer.empty_tag('pageSetUpPr', [ ['fitToPage', 1] ]) if fit_page?
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
      @writer.empty_tag('dimension', [ ['ref', ref] ])
    end
    #
    # Write the <sheetViews> element.
    #
    def write_sheet_views #:nodoc:
      @writer.tag_elements('sheetViews', []) { write_sheet_view }
    end

    def write_sheet_view #:nodoc:
      attributes = []
      # Hide screen gridlines if required
      attributes << ['showGridLines', 0] unless @screen_gridlines

      # Hide zeroes in cells.
      attributes << ['showZeros', 0] unless show_zeros?

      # Display worksheet right to left for Hebrew, Arabic and others.
      attributes << ['rightToLeft', 1] if @right_to_left

      # Show that the sheet tab is selected.
      attributes << ['tabSelected', 1] if @selected

      # Turn outlines off. Also required in the outlinePr element.
      attributes << ["showOutlineSymbols", 0] if @outline_on

      # Set the page view/layout mode if required.
      # TODO. Add pageBreakPreview mode when requested.
      attributes << ['view', 'pageLayout'] if page_view?

      # Set the zoom level.
      if @zoom != 100
        attributes << ['zoomScale', @zoom] unless page_view?
        attributes << ['zoomScaleNormal', @zoom] if zoom_scale_normal?
      end

      attributes << ['workbookViewId', 0]

      if @panes.empty? && @selections.empty?
        @writer.empty_tag('sheetView', attributes)
      else
        @writer.tag_elements('sheetView', attributes) do
          write_panes
          write_selections
        end
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
      attributes << ['pane', pane]              if pane
      attributes << ['activeCell', active_cell] if active_cell
      attributes << ['sqref', sqref]            if sqref

      @writer.empty_tag('selection', attributes)
    end

    #
    # Write the <sheetFormatPr> element.
    #
    def write_sheet_format_pr #:nodoc:
      base_col_width     = 10

      attributes = [
                    ['defaultRowHeight', @default_row_height]
                   ]
      if @default_row_height != @original_row_height
        attributes << ['customHeight', 1]
      end

      if ptrue?(@default_row_zeroed)
        attributes << ['zeroHeight', 1]
      end

      attributes << ['outlineLevelRow', @outline_row_level] if @outline_row_level > 0
      attributes << ['outlineLevelCol', @outline_col_level] if @outline_col_level > 0
      if @excel_version == 2010
        attributes << ['x14ac:dyDescent', '0.25']
      end
      @writer.empty_tag('sheetFormatPr', attributes)
    end

    #
    # Write the <cols> element and <col> sub elements.
    #
    def write_cols #:nodoc:
      # Exit unless some column have been formatted.
      return if @colinfo.empty?

      @writer.tag_elements('cols') do
        @colinfo.keys.sort.each {|col| write_col_info(@colinfo[col]) }
      end
    end

    #
    # Write the <col> element.
    #
    def write_col_info(args) #:nodoc:
      @writer.empty_tag('col', col_info_attributes(args))
    end

    def col_info_attributes(args)
      min    = args[0] || 0     # First formatted column.
      max    = args[1] || 0     # Last formatted column.
      width  = args[2]          # Col width in user units.
      format = args[3]          # Format index.
      hidden = args[4] || 0     # Hidden flag.
      level  = args[5] || 0     # Outline level.
      collapsed = args[6] || 0  # Outline level.
      xf_index = format ? format.get_xf_index : 0

      custom_width = true
      custom_width = false if width.nil? && hidden == 0
      custom_width = false if width == 8.43

      width = hidden == 0 ? 8.43 : 0 unless width

      # Convert column width from user units to character width.
      if width && width < 1
        width =
         ((width * (MAX_DIGIT_WIDTH + PADDING) + 0.5).to_i / MAX_DIGIT_WIDTH.to_f * 256).to_i / 256.0
      else
        width =
          (((width * MAX_DIGIT_WIDTH + 0.5).to_i + PADDING).to_i/ MAX_DIGIT_WIDTH.to_f * 256).to_i / 256.0
      end
      width = width.to_i if width - width.to_i == 0

      attributes = [
          ['min',   min + 1],
          ['max',   max + 1],
          ['width', width]
      ]

      attributes << ['style',        xf_index] if xf_index != 0
      attributes << ['hidden',       1]        if hidden != 0
      attributes << ['customWidth',  1]        if custom_width
      attributes << ['outlineLevel', level]    if level != 0
      attributes << ['collapsed',    1]        if collapsed != 0
      attributes
    end

    #
    # Write the <sheetData> element.
    #
    def write_sheet_data #:nodoc:
      if !@dim_rowmin
        # If the dimensions aren't defined then there is no data to write.
        @writer.empty_tag('sheetData')
      else
        @writer.tag_elements('sheetData') { write_rows }
      end
    end

    #
    # Write out the worksheet data as a series of rows and cells.
    #
    def write_rows #:nodoc:
      calculate_spans

      (@dim_rowmin .. @dim_rowmax).each do |row_num|
        # Skip row if it doesn't contain row formatting or cell data.
        next if not_contain_formatting_or_data?(row_num)

        span_index = row_num / 16
        span       = @row_spans[span_index]

        # Write the cells if the row contains data.
        if @cell_data_table[row_num]
          args = @set_rows[row_num] || []
          write_row_element(row_num, span, *args) do
            write_cell_column_dimension(row_num)
          end
        elsif @comments[row_num]
          write_empty_row(row_num, span, *(@set_rows[row_num]))
        else
          # Row attributes only.
          write_empty_row(row_num, span, *(@set_rows[row_num]))
        end
      end
    end

    def not_contain_formatting_or_data?(row_num) # :nodoc:
      !@set_rows[row_num] && !@cell_data_table[row_num] && !@comments.has_comment_in_row?(row_num)
    end

    def write_cell_column_dimension(row_num)  # :nodoc:
      (@dim_colmin .. @dim_colmax).each do |col_num|
        @cell_data_table[row_num][col_num].write_cell if @cell_data_table[row_num][col_num]
      end
    end

    #
    # Write the <row> element.
    #
    def write_row_element(*args)  # :nodoc:
      @writer.tag_elements('row', row_attributes(args)) do
        yield
      end
    end

    #
    # Write and empty <row> element, i.e., attributes only, no cell data.
    #
    def write_empty_row(*args) #:nodoc:
      @writer.empty_tag('row', row_attributes(args))
    end

    def row_attributes(args)
      r, spans, height, format, hidden, level, collapsed, empty_row = args
      height    ||= @default_row_height
      hidden    ||= 0
      level     ||= 0
      xf_index = format ? format.get_xf_index : 0

      attributes = [['r',  r + 1]]

      attributes << ['spans',        spans]    if spans
      attributes << ['s',            xf_index] if ptrue?(xf_index)
      attributes << ['customFormat', 1]        if ptrue?(format)
      attributes << ['ht',           height]   if height != @original_row_height
      attributes << ['hidden',       1]        if ptrue?(hidden)
      attributes << ['customHeight', 1]        if height != @original_row_height
      attributes << ['outlineLevel', level]    if ptrue?(level)
      attributes << ['collapsed',    1]        if ptrue?(collapsed)

      if @excel_version == 2010
        attributes << ['x14ac:dyDescent', '0.25']
      end
      attributes
    end

    #
    # Write the frozen or split <pane> elements.
    #
    def write_panes #:nodoc:
      return if @panes.empty?

      if @panes[4] == 2
        write_split_panes
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
      attributes << ['xSplit',      x_split] if x_split > 0
      attributes << ['ySplit',      y_split] if y_split > 0
      attributes << ['topLeftCell', top_left_cell]
      attributes << ['activePane',  active_pane]
      attributes << ['state',       state]

      @writer.empty_tag('pane', attributes)
    end

    #
    # Write the <pane> element for split panes.
    #
    # See also, implementers note for split_panes().
    #
    def write_split_panes #:nodoc:
      row, col, top_row, left_col = @panes
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
      attributes << ['xSplit', x_split] if x_split > 0
      attributes << ['ySplit', y_split] if y_split > 0
      attributes << ['topLeftCell', top_left_cell]
      attributes << ['activePane', active_pane] if has_selection

      @writer.empty_tag('pane', attributes)
    end

    #
    # Convert column width from user units to pane split width.
    #
    def calculate_x_split_width(width) #:nodoc:
      # Convert to pixels.
      if width < 1
        pixels = int(width * 12 + 0.5)
      else
        pixels = (width * MAX_DIGIT_WIDTH + 0.5).to_i + PADDING
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
      @writer.empty_tag('sheetCalcPr', [ ['fullCalcOnLoad', 1] ])
    end

    #
    # Write the <phoneticPr> element.
    #
    def write_phonetic_pr #:nodoc:
      attributes = [
                    ['fontId', 0],
                    ['type',   'noConversion']
                   ]

      @writer.empty_tag('phoneticPr', attributes)
    end

    #
    # Write the <pageMargins> element.
    #
    def write_page_margins #:nodoc:
      @page_setup.write_page_margins(@writer)
    end

    #
    # Write the <pageSetup> element.
    #
    def write_page_setup #:nodoc:
      @page_setup.write_page_setup(@writer)
    end

    #
    # Write the <mergeCells> element.
    #
    def write_merge_cells #:nodoc:
      write_some_elements('mergeCells', @merge) do
        @merge.each { |merged_range| write_merge_cell(merged_range) }
      end
    end

    def write_some_elements(tag, container)
      return if container.empty?

      @writer.tag_elements(tag, [ ['count', container.size] ]) do
        yield
      end
    end

    #
    # Write the <mergeCell> element.
    #
    def write_merge_cell(merged_range) #:nodoc:
      row_min, col_min, row_max, col_max = merged_range

      # Convert the merge dimensions to a cell range.
      cell_1 = xl_rowcol_to_cell(row_min, col_min)
      cell_2 = xl_rowcol_to_cell(row_max, col_max)

      @writer.empty_tag('mergeCell', [ ['ref', "#{cell_1}:#{cell_2}"] ])
    end

    #
    # Write the <printOptions> element.
    #
    def write_print_options #:nodoc:
      @page_setup.write_print_options(@writer)
    end

    #
    # Write the <headerFooter> element.
    #
    def write_header_footer #:nodoc:
      @page_setup.write_header_footer(@writer, excel2003_style?)
    end

    #
    # Write the <rowBreaks> element.
    #
    def write_row_breaks #:nodoc:
      write_breaks('rowBreaks')
    end

    #
    # Write the <colBreaks> element.
    #
    def write_col_breaks #:nodoc:
      write_breaks('colBreaks')
    end

    def write_breaks(tag) # :nodoc:
      case tag
      when 'rowBreaks'
        page_breaks = sort_pagebreaks(*(@page_setup.hbreaks))
        max = 16383
      when 'colBreaks'
        page_breaks = sort_pagebreaks(*(@page_setup.vbreaks))
        max = 1048575
      else
        raise "Invalid parameter '#{tag}' in write_breaks."
      end
      count = page_breaks.size

      return if page_breaks.empty?

      attributes = [
                    ['count', count],
                    ['manualBreakCount', count]
                   ]

      @writer.tag_elements(tag, attributes) do
        page_breaks.each { |num| write_brk(num, max) }
      end
    end
    #
    # Write the <brk> element.
    #
    def write_brk(id, max) #:nodoc:
      attributes = [
        ['id',  id],
        ['max', max],
        ['man', 1]
      ]

      @writer.empty_tag('brk', attributes)
    end

    #
    # Write the <autoFilter> element.
    #
    def write_auto_filter #:nodoc:
      return unless autofilter_ref?

      attributes = [
                    ['ref', @autofilter_ref]
                   ]

      if filter_on?
        # Autofilter defined active filters.
        @writer.tag_elements('autoFilter', attributes) do
          write_autofilters
        end
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

        # Filters are relative to first column in the autofilter.
        write_filter_column(col - col1, type, *tokens)
      end
    end

    #
    # Write the <filterColumn> element.
    #
    def write_filter_column(col_id, type, *filters) #:nodoc:
      @writer.tag_elements('filterColumn', [ ['colId', col_id] ]) do
        if type == 1
          # Type == 1 is the new XLSX style filter.
          write_filters(*filters)
        else
          # Type == 0 is the classic "custom" filter.
          write_custom_filters(*filters)
        end
      end
    end

    #
    # Write the <filters> element.
    #
    def write_filters(*filters) #:nodoc:
      if filters.size == 1 && filters[0] == 'blanks'
        # Special case for blank cells only.
        @writer.empty_tag('filters', [ ['blank', 1] ])
      else
        # General case.
        @writer.tag_elements('filters') do
          filters.each { |filter| write_filter(filter) }
        end
      end
    end

    #
    # Write the <filter> element.
    #
    def write_filter(val) #:nodoc:
      @writer.empty_tag('filter', [ ['val', val] ])
    end

    #
    # Write the <customFilters> element.
    #
    def write_custom_filters(*tokens) #:nodoc:
      if tokens.size == 2
        # One filter expression only.
        @writer.tag_elements('customFilters') { write_custom_filter(*tokens) }
      else
        # Two filter expressions.

        # Check if the "join" operand is "and" or "or".
        if tokens[2] == 0
          attributes = [ ['and', 1] ]
        else
          attributes = [ ['and', 0] ]
        end

        # Write the two custom filters.
        @writer.tag_elements('customFilters', attributes) do
          write_custom_filter(tokens[0], tokens[1])
          write_custom_filter(tokens[3], tokens[4])
        end
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
      attributes << ['operator', operator] unless operator == 'equal'
      attributes << ['val', val]

      @writer.empty_tag('customFilter', attributes)
    end

    #
    # Process any sored hyperlinks in row/col order and write the <hyperlinks>
    # element. The attributes are different for internal and external links.
    #
    def write_hyperlinks #:nodoc:
      return unless @hyperlinks
      hlink_attributes = []
      @hyperlinks.keys.sort.each do |row_num|
        # Sort the hyperlinks into column order.
        col_nums = @hyperlinks[row_num].keys.sort
        # Iterate over the columns.
        col_nums.each do |col_num|
          # Get the link data for this cell.
          link = @hyperlinks[row_num][col_num]

          # If the cell isn't a string then we have to add the url as
          # the string to display
          if ptrue?(@cell_data_table)          &&
              ptrue?(@cell_data_table[row_num]) &&
              ptrue?(@cell_data_table[row_num][col_num])
            if @cell_data_table[row_num][col_num].display_url_string?
              link.display_on
            end
          end

          if link.respond_to?(:external_hyper_link)
            # External link with rel file relationship.
            @rel_count += 1
            # Links for use by the packager.
            @external_hyper_links << link.external_hyper_link
          end
          hlink_attributes << link.attributes(row_num, col_num, @rel_count)
        end
      end

      return if hlink_attributes.empty?

      # Write the hyperlink elements.
      @writer.tag_elements('hyperlinks') do
        hlink_attributes.each do |attributes|
          @writer.empty_tag('hyperlink', attributes)
        end
      end
    end

    #
    # Write the <tabColor> element.
    #
    def write_tab_color #:nodoc:
      return unless tab_color?

      @writer.empty_tag('tabColor',
                        [
                         ['rgb', palette_color(@tab_color)]
                        ])
    end

    #
    # Write the <outlinePr> element.
    #
    def write_outline_pr
      return unless outline_changed?

      attributes = []
      attributes << ["applyStyles",  1] if @outline_style != 0
      attributes << ["summaryBelow", 0] if @outline_below == 0
      attributes << ["summaryRight", 0] if @outline_right == 0
      attributes << ["showOutlineSymbols", 0] if @outline_on == 0

      @writer.empty_tag('outlinePr', attributes)
    end

    #
    # Write the <sheetProtection> element.
    #
    def write_sheet_protection #:nodoc:
      return unless protect?

      attributes = []
      attributes << ["password",         @protect[:password]] if ptrue?(@protect[:password])
      attributes << ["sheet",            1] if ptrue?(@protect[:sheet])
      attributes << ["content",          1] if ptrue?(@protect[:content])
      attributes << ["objects",          1] unless ptrue?(@protect[:objects])
      attributes << ["scenarios",        1] unless ptrue?(@protect[:scenarios])
      attributes << ["formatCells",      0] if ptrue?(@protect[:format_cells])
      attributes << ["formatColumns",    0] if ptrue?(@protect[:format_columns])
      attributes << ["formatRows",       0] if ptrue?(@protect[:format_rows])
      attributes << ["insertColumns",    0] if ptrue?(@protect[:insert_columns])
      attributes << ["insertRows",       0] if ptrue?(@protect[:insert_rows])
      attributes << ["insertHyperlinks", 0] if ptrue?(@protect[:insert_hyperlinks])
      attributes << ["deleteColumns",    0] if ptrue?(@protect[:delete_columns])
      attributes << ["deleteRows",       0] if ptrue?(@protect[:delete_rows])

      attributes << ["selectLockedCells", 1] unless ptrue?(@protect[:select_locked_cells])

      attributes << ["sort",        0] if ptrue?(@protect[:sort])
      attributes << ["autoFilter",  0] if ptrue?(@protect[:autofilter])
      attributes << ["pivotTables", 0] if ptrue?(@protect[:pivot_tables])

      attributes << ["selectUnlockedCells", 1] unless ptrue?(@protect[:select_unlocked_cells])

      @writer.empty_tag('sheetProtection', attributes)
    end

    #
    # Write the <drawing> elements.
    #
    def write_drawings #:nodoc:
      increment_rel_id_and_write_r_id('drawing') if drawing?
    end

    #
    # Write the <legacyDrawing> element.
    #
    def write_legacy_drawing #:nodoc:
      increment_rel_id_and_write_r_id('legacyDrawing') if has_vml?
    end

    #
    # Write the <legacyDrawingHF> element.
    #
    def write_legacy_drawing_hf # :nodoc:
      return unless has_header_vml?

      # Increment the relationship id for any drawings or comments.
      @rel_count += 1

      attributes = [ ['r:id', "rId#{@rel_count}"] ]
      @writer.empty_tag('legacyDrawingHF', attributes)
    end

    #
    # Write the underline font element.
    #
    def write_underline(writer, underline) #:nodoc:
      writer.empty_tag('u', underline_attributes(underline))
    end

    #
    # Write the <tableParts> element.
    #
    def write_table_parts
      return if @tables.empty?

      @writer.tag_elements('tableParts', [ ['count', tables_count] ]) do
        tables_count.times { increment_rel_id_and_write_r_id('tablePart') }
      end
    end

    #
    # Write the <tablePart> element.
    #
    def write_table_part(id)
      @writer.empty_tag('tablePart', [r_id_attributes(id)])
    end

    def increment_rel_id_and_write_r_id(tag)
      @rel_count += 1
      write_r_id(tag, @rel_count)
    end

    def write_r_id(tag, id)
      @writer.empty_tag(tag, [r_id_attributes(id)])
    end

    #
    # Write the <extLst> element and sparkline subelements.
    #
    def write_ext_sparklines  # :nodoc:
      @writer.tag_elements('extLst') { write_ext } unless @sparklines.empty?
    end

    def write_ext
      @writer.tag_elements('ext', write_ext_attributes) do
        write_sparkline_groups
      end
    end

    def write_ext_attributes
      [
       ['xmlns:x14', "#{OFFICE_URL}spreadsheetml/2009/9/main"],
       ['uri',       '{05C60535-1F16-4fd2-B633-F4F36F0B64E0}']
      ]
    end

    def write_sparkline_groups
      # Write the x14:sparklineGroups element.
      @writer.tag_elements('x14:sparklineGroups', sparkline_groups_attributes) do
        # Write the sparkline elements.
        @sparklines.reverse.each do |sparkline|
          sparkline.write_sparkline_group(@writer)
        end
      end
    end

    def sparkline_groups_attributes  # :nodoc:
      [
       ['xmlns:xm', "#{OFFICE_URL}excel/2006/main"]
      ]
    end

    #
    # Write the <dataValidations> element.
    #
    def write_data_validations #:nodoc:
      write_some_elements('dataValidations', @validations) do
        @validations.each { |validation| validation.write_data_validation(@writer) }
      end
    end

    #
    # Write the Worksheet conditional formats.
    #
    def write_conditional_formats  #:nodoc:
      @cond_formats.keys.sort.each do |range|
        write_conditional_formatting(range, @cond_formats[range])
      end
    end

    #
    # Write the <conditionalFormatting> element.
    #
    def write_conditional_formatting(range, cond_formats) #:nodoc:
      @writer.tag_elements('conditionalFormatting', [ ['sqref', range] ]) do
        cond_formats.each { |cond_format| cond_format.write_cf_rule }
      end
    end

    def store_data_to_table(cell_data) #:nodoc:
      row, col = cell_data.row, cell_data.col
      if @cell_data_table[row]
        @cell_data_table[row][col] = cell_data
      else
        @cell_data_table[row] = {}
        @cell_data_table[row][col] = cell_data
      end
    end

    def store_row_col_max_min_values(row, col)
      store_row_max_min_values(row)
      store_col_max_min_values(col)
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
        if @cell_data_table[row_num]
          span_min, span_max = calc_spans(@cell_data_table, row_num, span_min, span_max)
        end

        # Calculate spans for comments.
        if @comments[row_num]
          span_min, span_max = calc_spans(@comments, row_num, span_min, span_max)
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

    def calc_spans(data, row_num, span_min, span_max)
      (@dim_colmin .. @dim_colmax).each do |col_num|
        if data[row_num][col_num]
          if !span_min
            span_min = col_num
            span_max = col_num
          else
            span_min = col_num if col_num < span_min
            span_max = col_num if col_num > span_max
          end
        end
      end
      [span_min, span_max]
    end

    #
    # Add a string to the shared string table, if it isn't already there, and
    # return the string index.
    #
    def shared_string_index(str, params = {}) #:nodoc:
      @workbook.shared_string_index(str, params)
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
      if row_num_1 == 0 and row_num_2 == ROW_MAX - 1
        range1       = col_char_1
        range2       = col_char_2
        row_col_only = true
      elsif col_num_1 == 0 and col_num_2 == COL_MAX - 1
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
      "#{quote_sheetname(@name)}!#{area}"
    end

    def fit_page? #:nodoc:
      @page_setup.fit_page
    end

    def filter_on? #:nodoc:
      ptrue?(@filter_on)
    end

    def tab_color? #:nodoc:
      ptrue?(@tab_color)
    end

    def outline_changed?
      ptrue?(@outline_changed)
    end

    def vba_codename?
      ptrue?(@vba_codename)
    end

    def zoom_scale_normal? #:nodoc:
      ptrue?(@zoom_scale_normal)
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

    def protect? #:nodoc:
      !!@protect
    end

    def autofilter_ref? #:nodoc:
      !!@autofilter_ref
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
        raise "Invalid column '#{col_letter}'" if col >= COL_MAX
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
