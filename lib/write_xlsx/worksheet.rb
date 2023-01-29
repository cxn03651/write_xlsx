# -*- coding: utf-8 -*-
# frozen_string_literal: true

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
  class Worksheet
    include Writexlsx::Utility

    MAX_DIGIT_WIDTH = 7    # For Calabri 11.  # :nodoc:
    PADDING         = 5                       # :nodoc:

    attr_reader :index                                            # :nodoc:
    attr_reader :charts, :images, :tables, :shapes, :drawings     # :nodoc:
    attr_reader :header_images, :footer_images, :background_image # :nodoc:
    attr_reader :vml_drawing_links                                # :nodoc:
    attr_reader :vml_data_id                                      # :nodoc:
    attr_reader :vml_header_id                                    # :nodoc:
    attr_reader :autofilter_area                                  # :nodoc:
    attr_reader :writer, :set_rows, :col_formats                  # :nodoc:
    attr_reader :vml_shape_id                                     # :nodoc:
    attr_reader :comments, :comments_author                       # :nodoc:
    attr_accessor :data_bars_2010, :dxf_priority                  # :nodoc:
    attr_reader :vba_codename                                     # :nodoc:
    attr_writer :excel_version

    def initialize(workbook, index, name) # :nodoc:
      @writer = Package::XMLWriterSimple.new

      @workbook = workbook
      @index = index
      @name = name
      @colinfo = {}
      @cell_data_table = []
      @excel_version = 2007
      @palette = workbook.palette
      @default_url_format = workbook.default_url_format
      @max_url_length = workbook.max_url_length

      @page_setup = PageSetup.new

      @screen_gridlines     = true
      @show_zeros           = true
      @dim_rowmin           = nil
      @dim_rowmax           = nil
      @dim_colmin           = nil
      @dim_colmax           = nil
      @selections           = []
      @panes                = []
      @hide_row_col_headers = 0

      @tab_color  = 0

      @set_cols = {}
      @set_rows = {}
      @zoom = 100
      @zoom_scale_normal = true
      @right_to_left = false
      @leading_zeros = false

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
      @external_background_links = []
      @drawing_links          = []
      @vml_drawing_links      = []
      @charts                 = []
      @images                 = []
      @tables                 = []
      @sparklines             = []
      @shapes                 = []
      @shape_hash             = {}
      @drawing_rels           = {}
      @drawing_rels_id        = 0
      @vml_drawing_rels       = {}
      @vml_drawing_rels_id    = 0
      @has_dynamic_arrays     = false
      @header_images          = []
      @footer_images          = []
      @background_image       = ''

      @outline_row_level = 0
      @outline_col_level = 0

      @original_row_height    = 15
      @default_row_height     = 15
      @default_row_pixels     = 20
      @default_col_width      = 8.43
      @default_col_pixels     = 64
      @default_row_rezoed     = 0

      @merge = []

      @has_vml        = false
      @has_header_vml = false
      @comments = Package::Comments.new(self)
      @buttons_array          = []
      @header_images_array    = []
      @ignore_errors          = nil

      @validations = []

      @cond_formats   = {}
      @data_bars_2010 = []
      @dxf_priority   = 1

      @protected_ranges     = []
      @num_protected_ranges = 0

      if excel2003_style?
        @original_row_height      = 12.75
        @default_row_height       = 12.75
        @default_row_pixels       = 17
        self.margins_left_right  = 0.75
        self.margins_top_bottom  = 1
        @page_setup.margin_header = 0.5
        @page_setup.margin_footer = 0.5
        @page_setup.header_footer_aligns = false
      end
    end

    def set_xml_writer(filename) # :nodoc:
      @writer.set_xml_writer(filename)
    end

    def assemble_xml_file # :nodoc:
      write_xml_declaration do
        @writer.tag_elements('worksheet', write_worksheet_attributes) do
          write_sheet_pr
          write_dimension
          write_sheet_views
          write_sheet_format_pr
          write_cols
          write_sheet_data
          write_sheet_protection
          write_protected_ranges
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
          write_ignored_errors
          write_drawings
          write_legacy_drawing
          write_legacy_drawing_hf
          write_picture
          write_table_parts
          write_ext_list
        end
      end
    end

    #
    # The name method is used to retrieve the name of a worksheet.
    #
    attr_reader :name

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
    def set_first_sheet
      @hidden = false
      @workbook.firstsheet = @index
    end

    #
    # Set the worksheet protection flags to prevent modification of worksheet
    # objects.
    #
    def protect(password = nil, options = {})
      check_parameter(options, protect_default_settings.keys, 'protect')
      @protect = protect_default_settings.merge(options)

      # Set the password after the user defined values.
      if password && password != ''
        @protect[:password] =
          encode_password(password)
      end
    end

    #
    # Unprotect ranges within a protected worksheet.
    #
    def unprotect_range(range, range_name = nil, password = nil)
      if range.nil?
        raise "The range must be defined in unprotect_range())\n"
      else
        range = range.gsub(/\$/, "")
        range = range.sub(/^=/, "")
        @num_protected_ranges += 1
      end

      range_name ||= "Range#{@num_protected_ranges}"
      password   &&= encode_password(password)

      @protected_ranges << [range, range_name, password]
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
    def set_column(*args)
      # Check for a cell reference in A1 notation and substitute row and column
      if args[0].to_s =~ /^\D/
        _row1, firstcol, _row2, lastcol, *data = substitute_cellref(*args)
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
      width ||= @default_col_width

      (firstcol..lastcol).each do |col|
        @col_sizes[col]   = [width, hidden]
        @col_formats[col] = format if format
      end
    end

    #
    # Set the width (and properties) of a single column or a range of columns in
    # pixels rather than character units.
    #
    def set_column_pixels(*data)
      cell = data[0]

      # Check for a cell reference in A1 notation and substitute row and column
      if cell =~ /^\D/
        data = substitute_cellref(*data)

        # Returned values row1 and row2 aren't required here. Remove them.
        data.shift         # $row1
        data.delete_at(1)  # $row2
      end

      # Ensure at least $first_col, $last_col and $width
      return if data.size < 3

      first_col = data[0]
      last_col  = data[1]
      pixels    = data[2]
      format    = data[3]
      hidden    = data[4] || 0
      level     = data[5]

      width = pixels_to_width(pixels) if ptrue?(pixels)

      set_column(first_col, last_col, width, format, hidden, level)
    end

    #
    # :call-seq:
    #   set_selection(cell_or_cell_range)
    #
    # Set which cell or cells are selected in a worksheet.
    #
    def set_selection(*args)
      return if args.empty?

      row_first, col_first, row_last, col_last = row_col_notation(args)
      active_cell = xl_rowcol_to_cell(row_first, col_first)

      if row_last  # Range selection.
        # Swap last row/col for first row/col as necessary
        row_first, row_last = row_last, row_first if row_first > row_last
        col_first, col_last = col_last, col_first if col_first > col_last

        sqref = xl_range(row_first, row_last, col_first, col_last)
      else          # Single cell selection.
        sqref = active_cell
      end

      # Selection isn't set for cell A1.
      return if sqref == 'A1'

      @selections = [[nil, active_cell, sqref]]
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
    def freeze_panes(*args)
      return if args.empty?

      # Check for a cell reference in A1 notation and substitute row and column.
      row, col, top_row, left_col, type = row_col_notation(args)

      col      ||= 0
      top_row  ||= row
      left_col ||= col
      type     ||= 0

      @panes   = [row, col, top_row, left_col, type]
    end

    #
    # :call-seq:
    #   split_panes(y, x, top_row, left_col)
    #
    # Set panes and mark them as split.
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
    def paper=(paper_size)
      @page_setup.paper = paper_size
    end

    def set_paper(paper_size)
      put_deprecate_message("#{self}.set_paper")
      self.paper = paper_size
    end

    #
    # Set the page header caption and optional margin.
    #
    def set_header(string = '', margin = 0.3, options = {})
      raise 'Header string must be less than 255 characters' if string.length > 255

      # Replace the Excel placeholder &[Picture] with the internal &G.
      @page_setup.header = string.gsub(/&\[Picture\]/, '&G')

      @page_setup.header_footer_aligns = options[:align_with_margins] if options[:align_with_margins]

      @page_setup.header_footer_scales = options[:scale_with_doc] if options[:scale_with_doc]

      # Reset the array in case the function is called more than once.
      @header_images = []

      [
        [:image_left, 'LH'], [:image_center, 'CH'], [:image_right, 'RH']
      ].each do |p|
        @header_images << [options[p.first], p.last] if options[p.first]
      end

      # placeholeder /&G/ の数
      placeholder_count = @page_setup.header.scan(/&G/).count

      image_count = @header_images.count

      raise "Number of header image (#{image_count}) doesn't match placeholder count (#{placeholder_count}) in string: #{@page_setup.header}" if image_count != placeholder_count

      @has_header_vml = true if image_count > 0

      @page_setup.margin_header         = margin || 0.3
      @page_setup.header_footer_changed = true
    end

    #
    # Set the page footer caption and optional margin.
    #
    def set_footer(string = '', margin = 0.3, options = {})
      raise 'Footer string must be less than 255 characters' if string.length > 255

      @page_setup.footer = string.dup

      # Replace the Excel placeholder &[Picture] with the internal &G.
      @page_setup.footer = string.gsub(/&\[Picture\]/, '&G')

      @page_setup.header_footer_aligns = options[:align_with_margins] if options[:align_with_margins]

      @page_setup.header_footer_scales = options[:scale_with_doc] if options[:scale_with_doc]

      # Reset the array in case the function is called more than once.
      @footer_images = []

      [
        [:image_left, 'LF'], [:image_center, 'CF'], [:image_right, 'RF']
      ].each do |p|
        @footer_images << [options[p.first], p.last] if options[p.first]
      end

      # placeholeder /&G/ の数
      placeholder_count = @page_setup.footer.scan(/&G/).count

      image_count = @footer_images.count

      raise "Number of footer image (#{image_count}) doesn't match placeholder count (#{placeholder_count}) in string: #{@page_setup.footer}" if image_count != placeholder_count

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

      row1, col1, row2, col2 = row_col_notation(args)
      return if [row1, col1, row2, col2].include?(nil)

      # Ignore max print area since this is the same as no print area for Excel.
      return if row1 == 0 && col1 == 0 && row2 == ROW_MAX - 1 && col2 == COL_MAX - 1

      # Build up the print area range "=Sheet2!R1C1:R2C1"
      @page_setup.print_area = convert_name_area(row1, col1, row2, col2)
    end

    #
    # Set the worksheet zoom factor in the range <tt>10 <= scale <= 400</tt>:
    #
    def zoom=(scale)
      # Confine the scale to Excel's range
      @zoom = if scale < 10 or scale > 400
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
      @page_setup.black_white = true
    end

    #
    # Causes the write() method to treat integers with a leading zero as a string.
    # This ensures that any leading zeros such, as in zip codes, are maintained.
    #
    def keep_leading_zeros(flag = true)
      @leading_zeros = !!flag
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

    #
    # :call-seq:
    #  write(row, column [ , token [ , format ] ])
    #
    # Excel makes a distinction between data types such as strings, numbers,
    # blanks, formulas and hyperlinks. To simplify the process of writing
    # data the {#write()}[#method-i-write] method acts as a general alias for several more
    # specific methods:
    #
    def write(*args)
      # Check for a cell reference in A1 notation and substitute row and column
      row_col_args = row_col_notation(args)
      token = row_col_args[2] || ''
      token = token.to_s if token.instance_of?(Time)

      fmt = row_col_args[3]
      if fmt.respond_to?(:force_text_format?) && fmt.force_text_format?
        write_string(*args) # Force text format
      # Match an array ref.
      elsif token.respond_to?(:to_ary)
        write_row(*args)
      elsif token.respond_to?(:coerce)  # Numeric
        write_number(*args)
      # Match integer with leading zero(s)
      elsif @leading_zeros && token =~ /^0\d*$/
        write_string(*args)
      elsif token =~ /\A([+-]?)(?=\d|\.\d)\d*(\.\d*)?([Ee]([+-]?\d+))?\Z/
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
        if token =~ %r{\A[fh]tt?ps?://}
          write_url(*args)
        # Match mailto:
        elsif token =~ /\Amailto:/
          write_url(*args)
        # Match internal or external sheet link
        elsif token =~ /\A(?:in|ex)ternal:/
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
    def write_comment(*args)
      # Check for a cell reference in A1 notation and substitute row and column
      row, col, string, options = row_col_notation(args)
      raise WriteXLSXInsufficientArgumentError if [row, col, string].include?(nil)

      # Check that row and col are valid and store max and min values
      check_dimensions(row, col)
      store_row_col_max_min_values(row, col)

      @has_vml = true

      # Process the properties of the cell comment.
      @comments.add(@workbook, self, row, col, string, options)
    end

    #
    # :call-seq:
    #   write_number(row, column, number [ , format ] )
    #
    # Write an integer or a float to the cell specified by row and column:
    #
    def write_number(*args)
      # Check for a cell reference in A1 notation and substitute row and column
      row, col, num, xf = row_col_notation(args)
      raise WriteXLSXInsufficientArgumentError if row.nil? || col.nil? || num.nil?

      # Check that row and col are valid and store max and min values
      check_dimensions(row, col)
      store_row_col_max_min_values(row, col)

      store_data_to_table(NumberCellData.new(num, xf), row, col)
    end

    #
    # :call-seq:
    #   write_string(row, column, string [, format ] )
    #
    # Write a string to the specified row and column (zero indexed).
    # +format+ is optional.
    #
    def write_string(*args)
      # Check for a cell reference in A1 notation and substitute row and column
      row, col, str, xf = row_col_notation(args)
      str &&= str.to_s
      raise WriteXLSXInsufficientArgumentError if row.nil? || col.nil? || str.nil?

      # Check that row and col are valid and store max and min values
      check_dimensions(row, col)
      store_row_col_max_min_values(row, col)

      index = shared_string_index(str.length > STR_MAX ? str[0, STR_MAX] : str)

      store_data_to_table(StringCellData.new(index, xf), row, col)
    end

    #
    # :call-seq:
    #    write_rich_string(row, column, (string | format, string)+,  [,cell_format] )
    #
    # The write_rich_string() method is used to write strings with multiple formats.
    # The method receives string fragments prefixed by format objects. The final
    # format object is used as the cell format.
    #
    def write_rich_string(*args)
      # Check for a cell reference in A1 notation and substitute row and column
      row, col, *rich_strings = row_col_notation(args)
      raise WriteXLSXInsufficientArgumentError if [row, col, rich_strings[0]].include?(nil)

      xf = cell_format_of_rich_string(rich_strings)

      # Check that row and col are valid and store max and min values
      check_dimensions(row, col)
      store_row_col_max_min_values(row, col)

      fragments, _length = rich_strings_fragments(rich_strings)
      # can't allow 2 formats in a row
      return -4 unless fragments

      index = shared_string_index(xml_str_of_rich_string(fragments))

      store_data_to_table(StringCellData.new(index, xf), row, col)
    end

    #
    # :call-seq:
    #   write_blank(row, col, format)
    #
    # Write a blank cell to the specified row and column (zero indexed).
    # A blank cell is used to specify formatting without adding a string
    # or a number.
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

      store_data_to_table(BlankCellData.new(xf), row, col)
    end

    #
    # :call-seq:
    #   write_formula(row, column, formula [ , format [ , value ] ] )
    #
    # Write a formula or function to the cell specified by +row+ and +column+:
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

        store_data_to_table(FormulaCellData.new(formula, format, value), row, col)
      end
    end

    #
    # Internal method shared by the write_array_formula() and
    # write_dynamic_array_formula() methods.
    #
    def write_array_formula_base(type, *args)
      # Check for a cell reference in A1 notation and substitute row and column
      row1, col1, row2, col2, formula, xf, value = row_col_notation(args)
      raise WriteXLSXInsufficientArgumentError if [row1, col1, row2, col2, formula].include?(nil)

      # Swap last row/col with first row/col as necessary
      row1, row2 = row2, row1 if row1 > row2
      col1, col2 = col2, col1 if col1 > col2

      # Check that row and col are valid and store max and min values
      check_dimensions(row1, col1)
      check_dimensions(row2, col2)
      store_row_col_max_min_values(row1, col1)
      store_row_col_max_min_values(row2, col2)

      # Define array range
      range = if row1 == row2 && col1 == col2
                xl_rowcol_to_cell(row1, col1)
              else
                "#{xl_rowcol_to_cell(row1, col1)}:#{xl_rowcol_to_cell(row2, col2)}"
              end

      # Remove array formula braces and the leading =.
      formula = formula.sub(/^\{(.*)\}$/, '\1').sub(/^=/, '')

      store_data_to_table(
        if type == 'a'
          FormulaArrayCellData.new(formula, xf, range, value)
        elsif type == 'd'
          DynamicFormulaArrayCellData.new(formula, xf, range, value)
        else
          raise "invalid type in write_array_formula_base()."
        end,
        row1, col1
      )

      # Pad out the rest of the area with formatted zeroes.
      (row1..row2).each do |row|
        (col1..col2).each do |col|
          next if row == row1 && col == col1

          write_number(row, col, 0, xf)
        end
      end
    end

    #
    # write_array_formula(row1, col1, row2, col2, formula, format)
    #
    # Write an array formula to the specified row and column (zero indexed).
    #
    def write_array_formula(*args)
      write_array_formula_base('a', *args)
    end

    #
    # write_dynamic_array_formula(row1, col1, row2, col2, formula, format)
    #
    # Write a dynamic formula to the specified row and column (zero indexed).
    #
    def write_dynamic_array_formula(*args)
      write_array_formula_base('d', *args)
      @has_dynamic_arrays = true
    end

    #
    # write_boolean(row, col, val, format)
    #
    # Write a boolean value to the specified row and column (zero indexed).
    #
    def write_boolean(*args)
      row, col, val, xf = row_col_notation(args)
      raise WriteXLSXInsufficientArgumentError if row.nil? || col.nil?

      val = val ? 1 : 0  # Boolean value.
      # xf : cell format.

      # Check that row and col are valid and store max and min values
      check_dimensions(row, col)
      store_row_col_max_min_values(row, col)

      store_data_to_table(BooleanCellData.new(val, xf), row, col)
    end

    #
    # :call-seq:
    #   update_format_with_params(row, col, format_params)
    #
    # Update formatting of the cell to the specified row and column (zero indexed).
    #
    def update_format_with_params(*args)
      row, col, params = row_col_notation(args)
      raise WriteXLSXInsufficientArgumentError if row.nil? || col.nil? || params.nil?

      # Check that row and col are valid and store max and min values
      check_dimensions(row, col)
      store_row_col_max_min_values(row, col)

      format = nil
      cell_data = nil
      if @cell_data_table[row].nil? || @cell_data_table[row][col].nil?
        format = @workbook.add_format(params)
        write_blank(row, col, format)
      else
        if @cell_data_table[row][col].xf.nil?
          format = @workbook.add_format(params)
          cell_data = @cell_data_table[row][col]
        else
          format = @workbook.add_format
          cell_data = @cell_data_table[row][col]
          format.copy(cell_data.xf)
          format.set_format_properties(params)
        end
        # keep original value of cell
        value = if cell_data.is_a? FormulaCellData
                  "=#{cell_data.token}"
                elsif cell_data.is_a? FormulaArrayCellData
                  "{=#{cell_data.token}}"
                elsif cell_data.is_a? StringCellData
                  @workbook.shared_strings.string(cell_data.data[:sst_id])
                else
                  cell_data.data
                end
        write(row, col, value, format)
      end
    end

    #
    # :call-seq:
    #   update_range_format_with_params(row_first, col_first, row_last, col_last, format_params)
    #
    # Update formatting of cells in range to the specified row and column (zero indexed).
    #
    def update_range_format_with_params(*args)
      row_first, col_first, row_last, col_last, params = row_col_notation(args)

      raise WriteXLSXInsufficientArgumentError if [row_first, col_first, row_last, col_last, params].include?(nil)

      # Swap last row/col with first row/col as necessary
      row_first,  row_last = row_last,  row_first  if row_first > row_last
      col_first, col_last = col_last, col_first if col_first > col_last

      # Check that column number is valid and store the max value
      check_dimensions(row_last, col_last)
      store_row_col_max_min_values(row_last, col_last)

      (row_first..row_last).each do |row|
        (col_first..col_last).each do |col|
          update_format_with_params(row, col, params)
        end
      end
    end

    #
    # The outline_settings() method is used to control the appearance of
    # outlines in Excel.
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

      raise "URL '#{url}' added but URL exceeds Excel's limit of 65,530 URLs per worksheet." if hyperlinks_count > 65_530

      # Add the default URL format.
      xf ||= @default_url_format

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
    def write_date_time(*args)
      # Check for a cell reference in A1 notation and substitute row and column
      row, col, str, xf = row_col_notation(args)
      raise WriteXLSXInsufficientArgumentError if [row, col, str].include?(nil)

      # Check that row and col are valid and store max and min values
      check_dimensions(row, col)
      store_row_col_max_min_values(row, col)

      date_time = convert_date_time(str)

      if date_time
        store_data_to_table(NumberCellData.new(date_time, xf), row, col)
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
    def insert_chart(*args)
      # Check for a cell reference in A1 notation and substitute row and column.
      row, col, chart, *options = row_col_notation(args)
      raise WriteXLSXInsufficientArgumentError if [row, col, chart].include?(nil)

      if options.first.instance_of?(Hash)
        params = options.first
        x_offset = params[:x_offset]
        y_offset = params[:y_offset]
        x_scale  = params[:x_scale]
        y_scale  = params[:y_scale]
        anchor   = params[:object_position]

      else
        x_offset, y_offset, x_scale, y_scale, anchor = options
      end
      x_offset ||= 0
      y_offset ||= 0
      x_scale  ||= 1
      y_scale  ||= 1
      anchor   ||= 1

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

      @charts << [row, col, chart, x_offset, y_offset, x_scale, y_scale, anchor]
    end

    #
    # :call-seq:
    #   insert_image(row, column, filename, options)
    #
    def insert_image(*args)
      # Check for a cell reference in A1 notation and substitute row and column.
      row, col, image, *options = row_col_notation(args)
      raise WriteXLSXInsufficientArgumentError if [row, col, image].include?(nil)

      if options.first.instance_of?(Hash)
        # Newer hash bashed options
        params      = options.first
        x_offset    = params[:x_offset]
        y_offset    = params[:y_offset]
        x_scale     = params[:x_scale]
        y_scale     = params[:y_scale]
        anchor      = params[:object_position]
        url         = params[:url]
        tip         = params[:tip]
        description = params[:description]
        decorative  = params[:decorative]
      else
        x_offset, y_offset, x_scale, y_scale, anchor = options
      end
      x_offset ||= 0
      y_offset ||= 0
      x_scale  ||= 1
      y_scale  ||= 1
      anchor   ||= 2

      @images << [
        row, col, image, x_offset, y_offset,
        x_scale, y_scale, url, tip, anchor, description, decorative
      ]
    end

    #
    # :call-seq:
    #   repeat_formula(row, column, formula [ , format ] )
    #
    # Deprecated. This is a writeexcel gem's method that is no longer
    # required by WriteXLSX.
    #
    def repeat_formula(*args)
      # Check for a cell reference in A1 notation and substitute row and column.
      row, col, formula, format, *pairs = row_col_notation(args)
      raise WriteXLSXInsufficientArgumentError if [row, col].include?(nil)

      raise "Odd number of elements in pattern/replacement list" unless pairs.size.even?
      raise "Not a valid formula" unless formula.respond_to?(:to_ary)

      tokens  = formula.join("\t").split("\t")
      raise "No tokens in formula" if tokens.empty?

      value = nil
      if pairs[-2] == 'result'
        value = pairs.pop
        pairs.pop
      end
      until pairs.empty?
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

      # Store the row sizes for use when calculating image vertices.
      @row_sizes[row] = [height, hidden]
    end

    #
    # This method is used to set the height (in pixels) and the properties of the
    # row.
    #
    def set_row_pixels(*data)
      height = data[1]

      data[1] = pixels_to_height(height) if ptrue?(height)
      set_row(*data)
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

      @default_row_zeroed = 1 if ptrue?(zero_height)
    end

    #
    # merge_range(first_row, first_col, last_row, last_col, string, format)
    #
    # Merge a range of cells. The first cell should contain the data and the
    # others should be blank. All cells should contain the same format.
    #
    def merge_range(*args)
      row_first, col_first, row_last, col_last, string, format, *extra_args = row_col_notation(args)

      raise "Incorrect number of arguments" if [row_first, col_first, row_last, col_last, format].include?(nil)
      raise "Fifth parameter must be a format object" unless format.respond_to?(:xf_index)
      raise "Can't merge single cell" if row_first == row_last && col_first == col_last

      # Swap last row/col with first row/col as necessary
      row_first,  row_last = row_last,  row_first  if row_first > row_last
      col_first, col_last = col_last, col_first if col_first > col_last

      # Check that the data range is valid and store the max and min values.
      check_dimensions(row_first, col_first)
      check_dimensions(row_last,  col_last)
      store_row_col_max_min_values(row_first, col_first)
      store_row_col_max_min_values(row_last,  col_last)

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

      # Check that the data range is valid and store the max and min values.
      check_dimensions(row_first, col_first)
      check_dimensions(row_last,  col_last)
      store_row_col_max_min_values(row_first, col_first)
      store_row_col_max_min_values(row_last,  col_last)

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
    def insert_button(*args)
      @buttons_array << button_params(*row_col_notation(args))
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
    def data_validation(*args)
      validation = DataValidation.new(*args)
      @validations << validation unless validation.validate_none?
    end

    #
    # Set the option to hide gridlines on the screen and the printed page.
    #
    def hide_gridlines(option = 1)
      @screen_gridlines = (option != 2)

      @page_setup.hide_gridlines(option)
    end

    # Set the option to print the row and column headers on the printed page.
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
    # Set the option to hide the row and column headers in Excel.
    #
    def hide_row_col_headers
      @hide_row_col_headers = 1
    end

    #
    # The fit_to_pages() method is used to fit the printed area to a specific
    # number of pages both vertically and horizontally. If the printed area
    # exceeds the specified number of pages it will be scaled down to fit.
    # This guarantees that the printed area will always appear on the
    # specified number of pages even if the page size or margins change.
    #
    def fit_to_pages(width = 1, height = 1)
      @page_setup.fit_page   = true
      @page_setup.fit_width  = width
      @page_setup.fit_height = height
      @page_setup.page_setup_changed = true
    end

    #
    # :call-seq:
    #   autofilter(first_row, first_col, last_row, last_col)
    #
    # Set the autofilter area in the worksheet.
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
    def filter_column(col, expression)
      raise "Must call autofilter before filter_column" unless @autofilter_area

      col = prepare_filter_column(col)

      tokens = extract_filter_tokens(expression)

      raise "Incorrect number of tokens in expression '#{expression}'" unless tokens.size == 3 || tokens.size == 7

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
    def set_h_pagebreaks(*args)
      breaks = args.collect do |brk|
        Array(brk)
      end.flatten
      @page_setup.hbreaks += breaks
    end

    #
    # Store the vertical page breaks on a worksheet.
    #
    def set_v_pagebreaks(*args)
      @page_setup.vbreaks += args
    end

    #
    # This method is used to make all cell comments visible when a worksheet
    # is opened.
    #
    def show_comments(visible = true)
      @comments_visible = visible
    end

    #
    # This method is used to set the default author of all cell comments.
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

      row, col, chart, x_offset, y_offset, x_scale, y_scale, anchor = @charts[index]
      chart.id = chart_id - 1
      x_scale ||= 0
      y_scale ||= 0

      # Use user specified dimensions, if any.
      width  = chart.width  if ptrue?(chart.width)
      height = chart.height if ptrue?(chart.height)

      width  = (0.5 + (width  * x_scale)).to_i
      height = (0.5 + (height * y_scale)).to_i

      dimensions = position_object_emus(col, row, x_offset, y_offset, width, height, anchor)

      # Set the chart name for the embedded object if it has been specified.
      name = chart.name

      # Create a Drawing object to use with worksheet unless one already exists.
      drawing = Drawing.new(drawing_type, dimensions, 0, 0, name, nil, anchor, drawing_rel_index, 0, nil, 0)
      if drawings?
        @drawings.add_drawing_object(drawing)
      else
        @drawings = Drawings.new
        @drawings.add_drawing_object(drawing)
        @drawings.embedded = 1

        @external_drawing_links << ['/drawing', "../drawings/drawing#{drawing_id}.xml"]
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
      (row_start..row_end).each do |row_num|
        # Store nil if row doesn't exist.
        unless @cell_data_table[row_num]
          data << nil
          next
        end

        (col_start..col_end).each do |col_num|
          cell = @cell_data_table[row_num][col_num]
          if cell
            data << cell.data
          else
            data << nil
          end
        end
      end

      data
    end

    #
    # Calculate the vertices that define the position of a graphical object within
    # the worksheet in pixels.
    #
    def position_object_pixels(col_start, row_start, x1, y1, width, height, anchor = nil) # :nodoc:
      # Adjust start column for negative offsets.
      while x1 < 0 && col_start > 0
        x1 += size_col(col_start - 1)
        col_start -= 1
      end

      # Adjust start row for negative offsets.
      while y1 < 0 && row_start > 0
        y1 += size_row(row_start - 1)
        row_start -= 1
      end

      # Ensure that the image isn't shifted off the page at top left.
      x1 = 0 if x1 < 0
      y1 = 0 if y1 < 0

      # Calculate the absolute x offset of the top-left vertex.
      x_abs = if @col_size_changed
                (0..col_start - 1).inject(0) { |sum, col| sum += size_col(col, anchor) }
              else
                # Optimisation for when the column widths haven't changed.
                @default_col_pixels * col_start
              end
      x_abs += x1

      # Calculate the absolute y offset of the top-left vertex.
      # Store the column change to allow optimisations.
      y_abs = if @row_size_changed
                (0..row_start - 1).inject(0) { |sum, row| sum += size_row(row, anchor) }
              else
                # Optimisation for when the row heights haven't changed.
                @default_row_pixels * row_start
              end
      y_abs += y1

      # Adjust start column for offsets that are greater than the col width.
      while x1 >= size_col(col_start, anchor)
        x1 -= size_col(col_start)
        col_start += 1
      end

      # Adjust start row for offsets that are greater than the row height.
      while y1 >= size_row(row_start, anchor)
        y1 -= size_row(row_start)
        row_start += 1
      end

      # Initialise end cell to the same as the start cell.
      col_end = col_start
      row_end = row_start

      # Only offset the image in the cell if the row/col isn't hidden.
      width  += x1 if size_col(col_start, anchor) > 0
      height += y1 if size_row(row_start, anchor) > 0

      # Subtract the underlying cell widths to find the end cell of the object.
      while width >= size_col(col_end, anchor)
        width -= size_col(col_end, anchor)
        col_end += 1
      end

      # Subtract the underlying cell heights to find the end cell of the object.
      while height >= size_row(row_end, anchor)
        height -= size_row(row_end, anchor)
        row_end += 1
      end

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
    def write_cell_value(value = '') # :nodoc:
      return write_cell_formula('=NA()') if !value.nil? && value.is_a?(Float) && value.nan?

      value ||= ''
      value = value.to_i if value == value.to_i
      @writer.data_element('v', value)
    end

    #
    # Write the cell formula <f> element.
    #
    def write_cell_formula(formula = '') # :nodoc:
      @writer.data_element('f', formula)
    end

    #
    # Write the cell array formula <f> element.
    #
    def write_cell_array_formula(formula, range) # :nodoc:
      @writer.data_element(
        'f', formula,
        [
          %w[t array],
          ['ref', range]
        ]
      )
    end

    def date_1904? # :nodoc:
      @workbook.date_1904?
    end

    def excel2003_style? # :nodoc:
      @workbook.excel2003_style
    end

    #
    # Convert from an Excel internal colour index to a XML style #RRGGBB index
    # based on the default or user defined values in the Workbook palette.
    #
    def palette_color(index) # :nodoc:
      if index.to_s =~ /^#([0-9A-F]{6})$/i
        "FF#{::Regexp.last_match(1).upcase}"
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
        @external_background_links,
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
      (1..num_comments_block).each do |i|
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
    def prepare_tables(table_id, seen)
      if tables_count > 0
        id = table_id
        tables.each do |table|
          table.prepare(id)

          if seen[table.name]
            raise "error: invalid duplicate table name '#{table.name}' found."
          else
            seen[table.name] = 1
          end

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
      @vba_codename = vba_codename || @name
    end

    #
    # Ignore worksheet errors/warnings in user defined ranges.
    #
    def ignore_errors(ignores)
      # List of valid input parameters.
      valid_parameter_keys = %i[
        number_stored_as_text
        eval_error
        formula_differs
        formula_range
        formula_unlocked
        empty_cell_reference
        list_data_validation
        calculated_column
        two_digit_text_year
      ]

      raise "Unknown parameter '#{ignores.key - valid_parameter_keys}' in ignore_errors()." unless (ignores.keys - valid_parameter_keys).empty?

      @ignore_errors = ignores
    end

    def write_ext(url, &block)
      attributes = [
        ['xmlns:x14', "#{OFFICE_URL}spreadsheetml/2009/9/main"],
        ['uri',       url]
      ]
      @writer.tag_elements('ext', attributes, &block)
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

    def has_dynamic_arrays?
      @has_dynamic_arrays
    end

    private

    #
    # Get the index used to address a drawing rel link.
    #
    def drawing_rel_index(target = nil)
      if !target
        # Undefined values for drawings like charts will always be unique.
        @drawing_rels_id += 1
      elsif ptrue?(@drawing_rels[target])
        @drawing_rels[target]
      else
        @drawing_rels_id += 1
        @drawing_rels[target] = @drawing_rels_id
      end
    end

    #
    # Get the index used to address a vml_drawing rel link.
    #
    def get_vml_drawing_rel_index(target)
      if @vml_drawing_rels[target]
        @vml_drawing_rels[target]
      else
        @vml_drawing_rels_id += 1
        @vml_drawing_rels[target] = @vml_drawing_rels_id
      end
    end

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
      rich_strings.pop if rich_strings[-1].respond_to?(:xf_index)
    end

    #
    # Convert the list of format, string tokens to pairs of (format, string)
    # except for the first string fragment which doesn't require a default
    # formatting run. Use the default for strings without a leading format.
    #
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
          if last == 'format'
            # If previous token was a format just add the string.
            fragments << token
          else
            # If previous token wasn't a format add one before the string.
            fragments << default << token
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
      writer.start_tag('r') unless fragments[0].respond_to?(:xf_index)

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
      (row_first..row_last).each do |row|
        (col_first..col_last).each do |col|
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
    def extract_filter_tokens(expression = nil) # :nodoc:
      return [] unless expression

      tokens = []
      str = expression
      while str =~ /"(?:[^"]|"")*"|\S+/
        tokens << ::Regexp.last_match(0)
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
    def parse_filter_expression(expression, tokens) # :nodoc:
      # The number of tokens will be either 3 (for 1 expression)
      # or 7 (for 2  expressions).
      #
      if tokens.size == 7
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
    def parse_filter_tokens(expression, tokens)     # :nodoc:
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
        '>=' => 6
      }

      operator = operators[tokens[1]]
      token    = tokens[2]

      # Special handling of "Top" filter expressions.
      if tokens[0] =~ /^top|bottom$/i
        value = tokens[1]
        if value.to_s =~ /\D/ or value.to_i < 1 or value.to_i > 500
          raise "The value '#{value}' in expression '#{expression}' " +
                "must be in the range 1 to 500"
        end
        token.downcase!
        if token != 'items' and token != '%'
          raise "The type '#{token}' in expression '#{expression}' " +
                "must be either 'items' or '%'"
        end

        operator = if tokens[0] =~ /^top$/i
                     30
                   else
                     32
                   end

        operator += 1 if tokens[2] == '%'

        token    = value
      end

      if !operator and tokens[0]
        raise "Token '#{tokens[1]}' is not a valid operator " +
              "in filter expression '#{expression}'"
      end

      # Special handling for Blanks/NonBlanks.
      if token.to_s =~ /^blanks|nonblanks$/i
        # Only allow Equals or NotEqual in this context.
        if operator != 2 and operator != 5
          raise "The operator '#{tokens[1]}' in expression '#{expression}' " +
                "is not valid in relation to Blanks/NonBlanks'"
        end

        token.downcase!

        # The operator should always be 2 (=) to flag a "simple" equality in
        # the binary record. Therefore we convert <> to =.
        if token == 'blanks'
          token = ' ' if operator == 5
        elsif operator == 5
          operator = 2
          token    = 'blanks'
        else
          operator = 5
          token    = ' '
        end
      end

      # if the string token contains an Excel match character then change the
      # operator type to indicate a non "simple" equality.
      operator = 22 if operator == 2 and token.to_s =~ /[*?]/

      [operator, token]
    end

    #
    # This is an internal method that is used to filter elements of the array of
    # pagebreaks used in the _store_hbreak() and _store_vbreak() methods. It:
    #   1. Removes duplicate entries from the list.
    #   2. Sorts the list.
    #   3. Removes 0 from the list if present.
    #
    def sort_pagebreaks(*args) # :nodoc:
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
    # the worksheet in EMUs.
    #
    def position_object_emus(col_start, row_start, x1, y1, width, height, anchor = nil) # :nodoc:
      col_start, row_start, x1, y1, col_end, row_end, x2, y2, x_abs, y_abs =
        position_object_pixels(col_start, row_start, x1, y1, width, height, anchor)

      # Convert the pixel values to EMUs. See above.
      x1    = (0.5 + (9_525 * x1)).to_i
      y1    = (0.5 + (9_525 * y1)).to_i
      x2    = (0.5 + (9_525 * x2)).to_i
      y2    = (0.5 + (9_525 * y2)).to_i
      x_abs = (0.5 + (9_525 * x_abs)).to_i
      y_abs = (0.5 + (9_525 * y_abs)).to_i

      [col_start, row_start, x1, y1, col_end, row_end, x2, y2, x_abs, y_abs]
    end

    #
    # Convert the width of a cell from user's units to pixels. Excel rounds the
    # column width to the nearest pixel. If the width hasn't been set by the user
    # we use the default value. A hidden column is treated as having a width of
    # zero unless it has the special "object_position" of 4 (size with cells).
    #
    def size_col(col, anchor = 0) # :nodoc:
      # Look up the cell value to see if it has been changed.
      if @col_sizes[col]
        width, hidden = @col_sizes[col]

        # Convert to pixels.
        pixels = if hidden == 1 && anchor != 4
                   0
                 elsif width < 1
                   ((width * (MAX_DIGIT_WIDTH + PADDING)) + 0.5).to_i
                 else
                   ((width * MAX_DIGIT_WIDTH) + 0.5).to_i + PADDING
                 end
      else
        pixels = @default_col_pixels
      end
      pixels
    end

    #
    # Convert the height of a cell from user's units to pixels. If the height
    # hasn't been set by the user we use the default value. A hidden row is
    # treated as having a height of zero unless it has the special
    # "object_position" of 4 (size with cells).
    #
    def size_row(row, anchor = 0) # :nodoc:
      # Look up the cell value to see if it has been changed
      if @row_sizes[row]
        height, hidden = @row_sizes[row]

        pixels = if hidden == 1 && anchor != 4
                   0
                 else
                   (4 / 3.0 * height).to_i
                 end
      else
        pixels = (4 / 3.0 * @default_row_height).to_i
      end
      pixels
    end

    #
    # Convert the width of a cell from pixels to character units.
    #
    def pixels_to_width(pixels)
      max_digit_width = 7.0
      padding         = 5.0

      if pixels <= 12
        pixels / (max_digit_width + padding)
      else
        (pixels - padding) / max_digit_width
      end
    end

    #
    # Convert the height of a cell from pixels to character units.
    #
    def pixels_to_height(pixels)
      height = 0.75 * pixels
      height = height.to_i if (height - height.to_i).abs < 0.1
      height
    end

    #
    # Set up image/drawings.
    #
    def prepare_image(index, image_id, drawing_id, width, height, name, image_type, x_dpi = 96, y_dpi = 96, md5 = nil) # :nodoc:
      x_dpi ||= 96
      y_dpi ||= 96
      drawing_type = 2

      row, col, _image, x_offset, y_offset,
      x_scale, y_scale, url, tip, anchor, description, decorative = @images[index]

      width  *= x_scale
      height *= y_scale

      width  *= 96.0 / x_dpi
      height *= 96.0 / y_dpi

      dimensions = position_object_emus(col, row, x_offset, y_offset, width, height, anchor)

      # Convert from pixels to emus.
      width  = (0.5 + (width  * 9_525)).to_i
      height = (0.5 + (height * 9_525)).to_i

      # Create a Drawing object to use with worksheet unless one already exists.
      drawing = Drawing.new(drawing_type, dimensions, width, height, name, nil, anchor, 0, 0, tip, decorative)
      if drawings?
        drawings = @drawings
      else
        drawings = Drawings.new
        drawings.embedded = 1

        @drawings = drawings

        @external_drawing_links << ['/drawing', "../drawings/drawing#{drawing_id}.xml"]
      end
      drawings.add_drawing_object(drawing)

      drawing.description = description if description

      if url
        rel_type = '/hyperlink'
        target_mode = 'External'
        target = escape_url(url) if url =~ %r{^[fh]tt?ps?://} || url =~ /^mailto:/
        if url =~ /^external:/
          target = escape_url(url.sub(/^external:/, ''))

          # Additional escape not required in worksheet hyperlinks
          target = target.gsub(/#/, '%23')

          # Prefix absolute paths (not relative) with file:///
          target = if target =~ /^\w:/ || target =~ /^\\\\/
                     "file:///#{target}"
                   else
                     target.gsub(/\\/, '/')
                   end
        end

        if url =~ /^internal:/
          target      = url.sub(/^internal:/, '#')
          target_mode = nil
        end

        if target.length > 255
          raise <<"EOS"
Ignoring URL #{target} where link or anchor > 255 characters since it exceeds Excel's limit for URLS. See LIMITATIONS section of the WriteXLSX documentation.
EOS
        end

        @drawing_links << [rel_type, target, target_mode] if target && !@drawing_rels[url]
        drawing.url_rel_index = drawing_rel_index(url)
      end

      @drawing_links << ['/image', "../media/image#{image_id}.#{image_type}"] unless @drawing_rels[md5]
      drawing.rel_index = drawing_rel_index(md5)
    end
    public :prepare_image

    def prepare_header_image(image_id, width, height, name, image_type, position, x_dpi, y_dpi, md5)
      # Strip the extension from the filename.
      body = name.dup
      body[/\.[^.]+$/, 0] = ''

      @vml_drawing_links << ['/image', "../media/image#{image_id}.#{image_type}"] unless @vml_drawing_rels[md5]

      ref_id = get_vml_drawing_rel_index(md5)
      @header_images_array << [width, height, body, position, x_dpi, y_dpi, ref_id]
    end
    public :prepare_header_image

    #
    # Set the background image for the worksheet.
    #
    def set_background(image)
      raise "Couldn't locate #{image}: $!" unless File.exist?(image)

      @background_image = image
    end
    public :set_background

    #
    # Set up an image without a drawing object for the background image.
    #
    def prepare_background(image_id, image_type)
      @external_background_links <<
        ['/image', "../media/image#{image_id}.#{image_type}"]
    end
    public :prepare_background

    #
    # :call-seq:
    #   insert_shape(row, col, shape [ , x, y, x_scale, y_scale ] )
    #
    # Insert a shape into the worksheet.
    #
    def insert_shape(*args)
      # Check for a cell reference in A1 notation and substitute row and column.
      row_start, column_start, shape, x_offset, y_offset, x_scale, y_scale, anchor =
        row_col_notation(args)
      raise "Insufficient arguments in insert_shape()" if [row_start, column_start, shape].include?(nil)

      shape.set_position(
        row_start, column_start, x_offset, y_offset,
        x_scale, y_scale, anchor
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

      insert = if ptrue?(shape.stencil)
                 # Insert a copy of the shape, not a reference so that the shape is
                 # used as a stencil. Previously stamped copies don't get modified
                 # if the stencil is modified.
                 shape.dup
               else
                 shape
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
      unless drawings?
        @drawings = Drawings.new
        @drawings.embedded = 1
        @external_drawing_links << ['/drawing', "../drawings/drawing#{drawing_id}.xml"]
        @has_shapes = true
      end

      # Validate the he shape against various rules.
      shape.validate(index)
      shape.calc_position_emus(self)

      drawing_type = 3
      drawing = Drawing.new(
        drawing_type, shape.dimensions, shape.width_emu, shape.height_emu,
        shape.name, shape, shape.anchor, drawing_rel_index, 0, nil, 0
      )
      drawings.add_drawing_object(drawing)
    end
    public :prepare_shape

    #
    # This method handles the parameters passed to insert_button as well as
    # calculating the button object position and vertices.
    #
    def button_params(row, col, params)
      button = Writexlsx::Package::Button.new

      button_number = 1 + @buttons_array.size

      # Set the button caption.
      caption = params[:caption] || "Button #{button_number}"

      button.font = { :_caption => caption }

      # Set the macro name.
      button.macro = if params[:macro]
                       "[0]!#{params[:macro]}"
                     else
                       "[0]!Button#{button_number}_Click"
                     end

      # Ensure that a width and height have been set.
      default_width  = @default_col_pixels
      default_height = @default_row_pixels
      params[:width]  = default_width  unless params[:width]
      params[:height] = default_height unless params[:height]

      # Set the x/y offsets.
      params[:x_offset] = 0 unless params[:x_offset]
      params[:y_offset] = 0 unless params[:y_offset]

      # Scale the size of the button box if required.
      params[:width] = params[:width] * params[:x_scale] if params[:x_scale]
      params[:height] = params[:height] * params[:y_scale] if params[:y_scale]

      # Round the dimensions to the nearest pixel.
      params[:width]  = (0.5 + params[:width]).to_i
      params[:height] = (0.5 + params[:height]).to_i

      params[:start_row] = row
      params[:start_col] = col

      # Calculate the positions of button object.
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
    def encode_password(password) # :nodoc:
      i = 0
      chars = password.split(//)
      count = chars.size

      chars.collect! do |char|
        i += 1
        char     = char.ord << i
        low_15   = char & 0x7fff
        high_15  = char & (0x7fff << 15)
        high_15  = high_15 >> 15
        char     = low_15 | high_15
      end

      encoded_password = 0x0000
      chars.each { |c| encoded_password ^= c }
      encoded_password ^= count
      encoded_password ^= 0xCE4B

      sprintf("%X", encoded_password)
    end

    #
    # Write the <worksheet> element. This is the root element of Worksheet.
    #
    def write_worksheet_attributes # :nodoc:
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
    def write_sheet_pr # :nodoc:
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
    def write_page_set_up_pr # :nodoc:
      @writer.empty_tag('pageSetUpPr', [['fitToPage', 1]]) if fit_page?
    end

    # Write the <dimension> element. This specifies the range of cells in the
    # worksheet. As a special case, empty spreadsheets use 'A1' as a range.
    #
    def write_dimension # :nodoc:
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
      @writer.empty_tag('dimension', [['ref', ref]])
    end

    #
    # Write the <sheetViews> element.
    #
    def write_sheet_views # :nodoc:
      @writer.tag_elements('sheetViews', []) { write_sheet_view }
    end

    def write_sheet_view # :nodoc:
      attributes = []
      # Hide screen gridlines if required.
      attributes << ['showGridLines', 0] unless @screen_gridlines

      # Hide the row/column headers.
      attributes << ['showRowColHeaders', 0] if ptrue?(@hide_row_col_headers)

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
      attributes << %w[view pageLayout] if page_view?

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
    def write_selections # :nodoc:
      @selections.each { |selection| write_selection(*selection) }
    end

    #
    # Write the <selection> element.
    #
    def write_selection(pane, active_cell, sqref) # :nodoc:
      attributes  = []
      attributes << ['pane', pane]              if pane
      attributes << ['activeCell', active_cell] if active_cell
      attributes << ['sqref', sqref]            if sqref

      @writer.empty_tag('selection', attributes)
    end

    #
    # Write the <sheetFormatPr> element.
    #
    def write_sheet_format_pr # :nodoc:
      attributes = [
        ['defaultRowHeight', @default_row_height]
      ]
      attributes << ['customHeight', 1] if @default_row_height != @original_row_height

      attributes << ['zeroHeight', 1] if ptrue?(@default_row_zeroed)

      attributes << ['outlineLevelRow', @outline_row_level] if @outline_row_level > 0
      attributes << ['outlineLevelCol', @outline_col_level] if @outline_col_level > 0
      attributes << ['x14ac:dyDescent', '0.25'] if @excel_version == 2010
      @writer.empty_tag('sheetFormatPr', attributes)
    end

    #
    # Write the <cols> element and <col> sub elements.
    #
    def write_cols # :nodoc:
      # Exit unless some column have been formatted.
      return if @colinfo.empty?

      @writer.tag_elements('cols') do
        @colinfo.keys.sort.each { |col| write_col_info(@colinfo[col]) }
      end
    end

    #
    # Write the <col> element.
    #
    def write_col_info(args) # :nodoc:
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

      width ||= hidden == 0 ? @default_col_width : 0

      # Convert column width from user units to character width.
      width = if width && width < 1
                (((width * (MAX_DIGIT_WIDTH + PADDING)) + 0.5).to_i / MAX_DIGIT_WIDTH.to_f * 256).to_i / 256.0
              else
                ((((width * MAX_DIGIT_WIDTH) + 0.5).to_i + PADDING).to_i / MAX_DIGIT_WIDTH.to_f * 256).to_i / 256.0
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
    def write_sheet_data # :nodoc:
      if @dim_rowmin
        @writer.tag_elements('sheetData') { write_rows }
      else
        # If the dimensions aren't defined then there is no data to write.
        @writer.empty_tag('sheetData')
      end
    end

    #
    # Write out the worksheet data as a series of rows and cells.
    #
    def write_rows # :nodoc:
      calculate_spans

      (@dim_rowmin..@dim_rowmax).each do |row_num|
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
      (@dim_colmin..@dim_colmax).each do |col_num|
        @cell_data_table[row_num][col_num].write_cell(self, row_num, col_num) if @cell_data_table[row_num][col_num]
      end
    end

    #
    # Write the <row> element.
    #
    def write_row_element(*args, &block)  # :nodoc:
      @writer.tag_elements('row', row_attributes(args), &block)
    end

    #
    # Write and empty <row> element, i.e., attributes only, no cell data.
    #
    def write_empty_row(*args) # :nodoc:
      @writer.empty_tag('row', row_attributes(args))
    end

    def row_attributes(args)
      r, spans, height, format, hidden, level, collapsed, _empty_row = args
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

      attributes << ['x14ac:dyDescent', '0.25'] if @excel_version == 2010
      attributes
    end

    #
    # Write the frozen or split <pane> elements.
    #
    def write_panes # :nodoc:
      return if @panes.empty?

      if @panes[4] == 2
        write_split_panes
      else
        write_freeze_panes(*@panes)
      end
    end

    #
    # Write the <pane> element for freeze panes.
    #
    def write_freeze_panes(row, col, top_row, left_col, type) # :nodoc:
      y_split       = row
      x_split       = col
      top_left_cell = xl_rowcol_to_cell(top_row, left_col)

      # Move user cell selection to the panes.
      unless @selections.empty?
        _dummy, active_cell, sqref = @selections[0]
        @selections = []
      end

      active_cell ||= nil
      sqref       ||= nil
      active_pane = set_active_pane_and_cell_selections(row, col, row, col, active_cell, sqref)

      # Set the pane type.
      state = if type == 0
                'frozen'
              elsif type == 1
                'frozenSplit'
              else
                'split'
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
    def write_split_panes # :nodoc:
      row, col, top_row, left_col = @panes
      has_selection = false
      y_split = row
      x_split = col

      # Move user cell selection to the panes.
      unless @selections.empty?
        _dummy, active_cell, sqref = @selections[0]
        @selections = []
        has_selection = true
      end

      # Convert the row and col to 1/20 twip units with padding.
      y_split = ((20 * y_split) + 300).to_i if y_split > 0
      x_split = calculate_x_split_width(x_split) if x_split > 0

      # For non-explicit topLeft definitions, estimate the cell offset based
      # on the pixels dimensions. This is only a workaround and doesn't take
      # adjusted cell dimensions into account.
      if top_row == row && left_col == col
        top_row  = (0.5 + ((y_split - 300) / 20 / 15)).to_i
        left_col = (0.5 + ((x_split - 390) / 20 / 3 * 4 / 64)).to_i
      end

      top_left_cell = xl_rowcol_to_cell(top_row, left_col)

      # If there is no selection set the active cell to the top left cell.
      unless has_selection
        active_cell = top_left_cell
        sqref       = top_left_cell
      end
      active_pane = set_active_pane_and_cell_selections(
        row, col, top_row, left_col, active_cell, sqref
      )

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
    def calculate_x_split_width(width) # :nodoc:
      # Convert to pixels.
      pixels = if width < 1
                 int((width * 12) + 0.5)
               else
                 ((width * MAX_DIGIT_WIDTH) + 0.5).to_i + PADDING
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
    def write_sheet_calc_pr # :nodoc:
      @writer.empty_tag('sheetCalcPr', [['fullCalcOnLoad', 1]])
    end

    #
    # Write the <phoneticPr> element.
    #
    def write_phonetic_pr # :nodoc:
      attributes = [
        ['fontId', 0],
        %w[type noConversion]
      ]

      @writer.empty_tag('phoneticPr', attributes)
    end

    #
    # Write the <pageMargins> element.
    #
    def write_page_margins # :nodoc:
      @page_setup.write_page_margins(@writer)
    end

    #
    # Write the <pageSetup> element.
    #
    def write_page_setup # :nodoc:
      @page_setup.write_page_setup(@writer)
    end

    #
    # Write the <mergeCells> element.
    #
    def write_merge_cells # :nodoc:
      write_some_elements('mergeCells', @merge) do
        @merge.each { |merged_range| write_merge_cell(merged_range) }
      end
    end

    def write_some_elements(tag, container, &block)
      return if container.empty?

      @writer.tag_elements(tag, [['count', container.size]], &block)
    end

    #
    # Write the <mergeCell> element.
    #
    def write_merge_cell(merged_range) # :nodoc:
      row_min, col_min, row_max, col_max = merged_range

      # Convert the merge dimensions to a cell range.
      cell_1 = xl_rowcol_to_cell(row_min, col_min)
      cell_2 = xl_rowcol_to_cell(row_max, col_max)

      @writer.empty_tag('mergeCell', [['ref', "#{cell_1}:#{cell_2}"]])
    end

    #
    # Write the <printOptions> element.
    #
    def write_print_options # :nodoc:
      @page_setup.write_print_options(@writer)
    end

    #
    # Write the <headerFooter> element.
    #
    def write_header_footer # :nodoc:
      @page_setup.write_header_footer(@writer, excel2003_style?)
    end

    #
    # Write the <rowBreaks> element.
    #
    def write_row_breaks # :nodoc:
      write_breaks('rowBreaks')
    end

    #
    # Write the <colBreaks> element.
    #
    def write_col_breaks # :nodoc:
      write_breaks('colBreaks')
    end

    def write_breaks(tag) # :nodoc:
      case tag
      when 'rowBreaks'
        page_breaks = sort_pagebreaks(*@page_setup.hbreaks)
        max = 16383
      when 'colBreaks'
        page_breaks = sort_pagebreaks(*@page_setup.vbreaks)
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
    def write_brk(id, max) # :nodoc:
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
    def write_auto_filter # :nodoc:
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
    def write_autofilters # :nodoc:
      col1, col2 = @filter_range

      (col1..col2).each do |col|
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
    def write_filter_column(col_id, type, *filters) # :nodoc:
      @writer.tag_elements('filterColumn', [['colId', col_id]]) do
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
    def write_filters(*filters) # :nodoc:
      non_blanks = filters.reject { |filter| filter.to_s =~ /^blanks$/i }
      attributes = []

      attributes = [['blank', 1]] if filters != non_blanks

      if filters.size == 1 && non_blanks.empty?
        # Special case for blank cells only.
        @writer.empty_tag('filters', attributes)
      else
        # General case.
        @writer.tag_elements('filters', attributes) do
          non_blanks.sort.each { |filter| write_filter(filter) }
        end
      end
    end

    #
    # Write the <filter> element.
    #
    def write_filter(val) # :nodoc:
      @writer.empty_tag('filter', [['val', val]])
    end

    #
    # Write the <customFilters> element.
    #
    def write_custom_filters(*tokens) # :nodoc:
      if tokens.size == 2
        # One filter expression only.
        @writer.tag_elements('customFilters') { write_custom_filter(*tokens) }
      else
        # Two filter expressions.

        # Check if the "join" operand is "and" or "or".
        attributes = if tokens[2] == 0
                       [['and', 1]]
                     else
                       [['and', 0]]
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
    def write_custom_filter(operator, val) # :nodoc:
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
    def write_hyperlinks # :nodoc:
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
             ptrue?(@cell_data_table[row_num][col_num]) && @cell_data_table[row_num][col_num].display_url_string?
            link.display_on
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
    def write_tab_color # :nodoc:
      return unless tab_color?

      @writer.empty_tag(
        'tabColor',
        [
          ['rgb', palette_color(@tab_color)]
        ]
      )
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
    def write_sheet_protection # :nodoc:
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
    # Write the <protectedRanges> element.
    #
    def write_protected_ranges
      return if @num_protected_ranges == 0

      @writer.tag_elements('protectedRanges') do
        @protected_ranges.each do |protected_range|
          write_protected_range(*protected_range)
        end
      end
    end

    #
    # Write the <protectedRange> element.
    #
    def write_protected_range(sqref, name, password)
      attributes = []

      attributes << ['password', password] if password
      attributes << ['sqref',    sqref]
      attributes << ['name',     name]

      @writer.empty_tag('protectedRange', attributes)
    end

    #
    # Write the <drawing> elements.
    #
    def write_drawings # :nodoc:
      increment_rel_id_and_write_r_id('drawing') if drawings?
    end

    #
    # Write the <legacyDrawing> element.
    #
    def write_legacy_drawing # :nodoc:
      increment_rel_id_and_write_r_id('legacyDrawing') if has_vml?
    end

    #
    # Write the <legacyDrawingHF> element.
    #
    def write_legacy_drawing_hf # :nodoc:
      return unless has_header_vml?

      # Increment the relationship id for any drawings or comments.
      @rel_count += 1

      attributes = [['r:id', "rId#{@rel_count}"]]
      @writer.empty_tag('legacyDrawingHF', attributes)
    end

    #
    # Write the <picture> element.
    #
    def write_picture
      return unless ptrue?(@background_image)

      # Increment the relationship id.
      @rel_count += 1
      id = @rel_count

      attributes = [['r:id', "rId#{id}"]]

      @writer.empty_tag('picture', attributes)
    end

    #
    # Write the underline font element.
    #
    def write_underline(writer, underline) # :nodoc:
      writer.empty_tag('u', underline_attributes(underline))
    end

    #
    # Write the <tableParts> element.
    #
    def write_table_parts
      return if @tables.empty?

      @writer.tag_elements('tableParts', [['count', tables_count]]) do
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
    # Write the <extLst> element for data bars and sparklines.
    #
    def write_ext_list  # :nodoc:
      return if @data_bars_2010.empty? && @sparklines.empty?

      @writer.tag_elements('extLst') do
        write_ext_list_data_bars  if @data_bars_2010.size > 0
        write_ext_list_sparklines if @sparklines.size > 0
      end
    end

    #
    # Write the Excel 2010 data_bar subelements.
    #
    def write_ext_list_data_bars
      # Write the ext element.
      write_ext('{78C0D931-6437-407d-A8EE-F0AAD7539E65}') do
        @writer.tag_elements('x14:conditionalFormattings') do
          # Write each of the Excel 2010 conditional formatting data bar elements.
          @data_bars_2010.each do |data_bar|
            # Write the x14:conditionalFormatting element.
            write_conditional_formatting_2010(data_bar)
          end
        end
      end
    end

    #
    # Write the <x14:conditionalFormatting> element.
    #
    def write_conditional_formatting_2010(data_bar)
      xmlns_xm = 'http://schemas.microsoft.com/office/excel/2006/main'

      attributes = [['xmlns:xm', xmlns_xm]]

      @writer.tag_elements('x14:conditionalFormatting', attributes) do
        # Write the '<x14:cfRule element.
        write_x14_cf_rule(data_bar)

        # Write the x14:dataBar element.
        write_x14_data_bar(data_bar)

        # Write the x14 max and min data bars.
        write_x14_cfvo(data_bar[:x14_min_type], data_bar[:min_value])
        write_x14_cfvo(data_bar[:x14_max_type], data_bar[:max_value])

        # Write the x14:borderColor element.
        write_x14_border_color(data_bar[:bar_border_color]) unless ptrue?(data_bar[:bar_no_border])

        # Write the x14:negativeFillColor element.
        write_x14_negative_fill_color(data_bar[:bar_negative_color]) unless ptrue?(data_bar[:bar_negative_color_same])

        # Write the x14:negativeBorderColor element.
        if !ptrue?(data_bar[:bar_no_border]) &&
           !ptrue?(data_bar[:bar_negative_border_color_same])
          write_x14_negative_border_color(
            data_bar[:bar_negative_border_color]
          )
        end

        # Write the x14:axisColor element.
        write_x14_axis_color(data_bar[:bar_axis_color]) if data_bar[:bar_axis_position] != 'none'

        # Write closing elements.
        @writer.end_tag('x14:dataBar')
        @writer.end_tag('x14:cfRule')

        # Add the conditional format range.
        @writer.data_element('xm:sqref', data_bar[:range])
      end
    end

    #
    # Write the <cfvo> element.
    #
    def write_x14_cfvo(type, value)
      attributes = [['type', type]]

      if %w[min max autoMin autoMax].include?(type)
        @writer.empty_tag('x14:cfvo', attributes)
      else
        @writer.tag_elements('x14:cfvo', attributes) do
          @writer.data_element('xm:f', value)
        end
      end
    end

    #
    # Write the <'<x14:cfRule> element.
    #
    def write_x14_cf_rule(data_bar)
      type = 'dataBar'
      id   = data_bar[:guid]

      attributes = [
        ['type', type],
        ['id',   id]
      ]

      @writer.start_tag('x14:cfRule', attributes)
    end

    #
    # Write the <x14:dataBar> element.
    #
    def write_x14_data_bar(data_bar)
      min_length = 0
      max_length = 100

      attributes = [
        ['minLength', min_length],
        ['maxLength', max_length]
      ]

      attributes << ['border',   1] unless ptrue?(data_bar[:bar_no_border])
      attributes << ['gradient', 0] if ptrue?(data_bar[:bar_solid])

      attributes << %w[direction leftToRight] if data_bar[:bar_direction] == 'left'
      attributes << %w[direction rightToLeft] if data_bar[:bar_direction] == 'right'

      attributes << ['negativeBarColorSameAsPositive', 1] if ptrue?(data_bar[:bar_negative_color_same])

      if !ptrue?(data_bar[:bar_no_border]) &&
         !ptrue?(data_bar[:bar_negative_border_color_same])
        attributes << ['negativeBarBorderColorSameAsPositive', 0]
      end

      attributes << %w[axisPosition middle] if data_bar[:bar_axis_position] == 'middle'

      attributes << %w[axisPosition none] if data_bar[:bar_axis_position] == 'none'

      @writer.start_tag('x14:dataBar', attributes)
    end

    #
    # Write the <x14:borderColor> element.
    #
    def write_x14_border_color(rgb)
      attributes = [['rgb', rgb]]

      @writer.empty_tag('x14:borderColor', attributes)
    end

    #
    # Write the <x14:negativeFillColor> element.
    #
    def write_x14_negative_fill_color(rgb)
      attributes = [['rgb', rgb]]

      @writer.empty_tag('x14:negativeFillColor', attributes)
    end

    #
    # Write the <x14:negativeBorderColor> element.
    #
    def write_x14_negative_border_color(rgb)
      attributes = [['rgb', rgb]]

      @writer.empty_tag('x14:negativeBorderColor', attributes)
    end

    #
    # Write the <x14:axisColor> element.
    #
    def write_x14_axis_color(rgb)
      attributes = [['rgb', rgb]]

      @writer.empty_tag('x14:axisColor', attributes)
    end

    #
    # Write the sparkline subelements.
    #
    def write_ext_list_sparklines
      # Write the ext element.
      write_ext('{05C60535-1F16-4fd2-B633-F4F36F0B64E0}') do
        # Write the x14:sparklineGroups element.
        write_sparkline_groups
      end
    end

    #
    # Write the <x14:sparklines> element and <x14:sparkline> subelements.
    #
    def write_sparklines(sparkline)
      # Write the sparkline elements.
      @writer.tag_elements('x14:sparklines') do
        (0..sparkline[:count] - 1).each do |i|
          range    = sparkline[:ranges][i]
          location = sparkline[:locations][i]

          @writer.tag_elements('x14:sparkline') do
            @writer.data_element('xm:f', range)
            @writer.data_element('xm:sqref', location)
          end
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
    def write_data_validations # :nodoc:
      write_some_elements('dataValidations', @validations) do
        @validations.each { |validation| validation.write_data_validation(@writer) }
      end
    end

    #
    # Write the Worksheet conditional formats.
    #
    def write_conditional_formats  # :nodoc:
      @cond_formats.keys.sort.each do |range|
        write_conditional_formatting(range, @cond_formats[range])
      end
    end

    #
    # Write the <conditionalFormatting> element.
    #
    def write_conditional_formatting(range, cond_formats) # :nodoc:
      @writer.tag_elements('conditionalFormatting', [['sqref', range]]) do
        cond_formats.each { |cond_format| cond_format.write_cf_rule }
      end
    end

    def store_data_to_table(cell_data, row, col) # :nodoc:
      if @cell_data_table[row]
        @cell_data_table[row][col] = cell_data
      else
        @cell_data_table[row] = []
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
    def calculate_spans # :nodoc:
      span_min = nil
      span_max = 0
      spans = []

      (@dim_rowmin..@dim_rowmax).each do |row_num|
        span_min, span_max = calc_spans(@cell_data_table, row_num, span_min, span_max) if @cell_data_table[row_num]

        # Calculate spans for comments.
        span_min, span_max = calc_spans(@comments, row_num, span_min, span_max) if @comments[row_num]

        next unless ((row_num + 1) % 16 == 0) || (row_num == @dim_rowmax)

        span_index = row_num / 16
        next unless span_min

        span_min += 1
        span_max += 1
        spans[span_index] = "#{span_min}:#{span_max}"
        span_min = nil
      end

      @row_spans = spans
    end

    def calc_spans(data, row_num, span_min, span_max)
      (@dim_colmin..@dim_colmax).each do |col_num|
        if data[row_num][col_num]
          if span_min
            span_min = col_num if col_num < span_min
            span_max = col_num if col_num > span_max
          else
            span_min = col_num
            span_max = col_num
          end
        end
      end
      [span_min, span_max]
    end

    #
    # Add a string to the shared string table, if it isn't already there, and
    # return the string index.
    #
    def shared_string_index(str) # :nodoc:
      @workbook.shared_string_index(str)
    end

    #
    # convert_name_area(first_row, first_col, last_row, last_col)
    #
    # Convert zero indexed rows and columns to the format required by worksheet
    # named ranges, eg, "Sheet1!$A$1:$C$13".
    #
    def convert_name_area(row_num_1, col_num_1, row_num_2, col_num_2) # :nodoc:
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
      area = if range1 == range2 && !row_col_only
               range1
             else
               "#{range1}:#{range2}"
             end

      # Build up the print area range "Sheet1!$A$1:$C$13".
      "#{quote_sheetname(@name)}!#{area}"
    end

    def fit_page? # :nodoc:
      @page_setup.fit_page
    end

    def filter_on? # :nodoc:
      ptrue?(@filter_on)
    end

    def tab_color? # :nodoc:
      ptrue?(@tab_color)
    end

    def outline_changed?
      ptrue?(@outline_changed)
    end

    def vba_codename?
      ptrue?(@vba_codename)
    end

    def zoom_scale_normal? # :nodoc:
      ptrue?(@zoom_scale_normal)
    end

    def page_view? # :nodoc:
      !!@page_view
    end

    def right_to_left? # :nodoc:
      !!@right_to_left
    end

    def show_zeros? # :nodoc:
      !!@show_zeros
    end

    def protect? # :nodoc:
      !!@protect
    end

    def autofilter_ref? # :nodoc:
      !!@autofilter_ref
    end

    def drawings? # :nodoc:
      !!@drawings
    end

    def remove_white_space(margin) # :nodoc:
      if margin.respond_to?(:gsub)
        margin.gsub(/[^\d.]/, '')
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
          ['topRight',    col_cell,    col_cell] <<
          ['bottomLeft',  row_cell,    row_cell] <<
          ['bottomRight', active_cell, sqref]
      elsif col > 0
        active_pane = 'topRight'
        @selections << ['topRight', active_cell, sqref]
      else
        active_pane = 'bottomLeft'
        @selections << ['bottomLeft', active_cell, sqref]
      end
      active_pane
    end

    def prepare_filter_column(col) # :nodoc:
      # Check for a column reference in A1 notation and substitute.
      if col.to_s =~ /^\D/
        col_letter = col

        # Convert col ref to a cell ref and then to a col number.
        _dummy, col = substitute_cellref("#{col}1")
        raise "Invalid column '#{col_letter}'" if col >= COL_MAX
      end

      col_first, col_last = @filter_range

      # Reject column if it is outside filter range.
      raise "Column '#{col}' outside autofilter column range (#{col_first} .. #{col_last})" if col < col_first or col > col_last

      col
    end

    #
    # Write the <ignoredErrors> element.
    #
    def write_ignored_errors
      return unless @ignore_errors

      ignore = @ignore_errors

      @writer.tag_elements('ignoredErrors') do
        {
          :number_stored_as_text => 'numberStoredAsText',
          :eval_error            => 'evalError',
          :formula_differs       => 'formula',
          :formula_range         => 'formulaRange',
          :formula_unlocked      => 'unlockedFormula',
          :empty_cell_reference  => 'emptyCellReference',
          :list_data_validation  => 'listDataValidation',
          :calculated_column     => 'calculatedColumn',
          :two_digit_text_year   => 'twoDigitTextYear'
        }.each do |key, value|
          write_ignored_error(value, ignore[key]) if ignore[key]
        end
      end
    end

    #
    # Write the <ignoredError> element.
    #
    def write_ignored_error(type, sqref)
      attributes = [
        ['sqref', sqref],
        [type, 1]
      ]

      @writer.empty_tag('ignoredError', attributes)
    end
  end
end
