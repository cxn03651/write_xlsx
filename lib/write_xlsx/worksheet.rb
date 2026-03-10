# -*- coding: utf-8 -*-
# frozen_string_literal: true

require 'write_xlsx/colors'
require 'write_xlsx/compatibility'
require 'write_xlsx/drawing'
require 'write_xlsx/format'
require 'write_xlsx/image'
require 'write_xlsx/image_property'
require 'write_xlsx/inserted_chart'
require 'write_xlsx/package/button'
require 'write_xlsx/package/conditional_format'
require 'write_xlsx/package/xml_writer_simple'
require 'write_xlsx/page_setup'
require 'write_xlsx/sparkline'
require 'write_xlsx/utility'
require 'write_xlsx/worksheet/asset_manager'
require 'write_xlsx/worksheet/cell_data'
require 'write_xlsx/worksheet/cell_data_manager'
require 'write_xlsx/worksheet/cell_data_store'
require 'write_xlsx/worksheet/data_validation'
require 'write_xlsx/worksheet/data_writing'
require 'write_xlsx/worksheet/drawing_methods'
require 'write_xlsx/worksheet/formatting'
require 'write_xlsx/worksheet/hyperlink'
require 'write_xlsx/worksheet/columns'
require 'write_xlsx/worksheet/rows'
require 'write_xlsx/worksheet/selection'
require 'write_xlsx/worksheet/panes'
require 'write_xlsx/worksheet/autofilter'
require 'write_xlsx/worksheet/conditional_formats'
require 'write_xlsx/worksheet/protection'
require 'write_xlsx/worksheet/print_options'
require 'write_xlsx/worksheet/xml_writer'
require 'forwardable'
require 'tempfile'
require 'date'

module Writexlsx
  class Worksheet
    extend Forwardable

    include Writexlsx::Utility
    include Autofilter
    include CellDataManager
    include Columns
    include ConditionalFormats
    include DrawingMethods
    include DataWriting
    include Formatting
    include Rows
    include Selection
    include Panes
    include Protection
    include PrintOptions
    include XmlWriter

    COLINFO = Struct.new('ColInfo', :width, :format, :hidden, :level, :collapsed, :autofit)

    attr_reader :index, :name                                     # :nodoc:
    attr_reader :drawings                                         # :nodoc:
    attr_reader :vml_drawing_links                                # :nodoc:
    attr_reader :vml_data_id                                      # :nodoc:
    attr_reader :vml_header_id                                    # :nodoc:
    attr_reader :autofilter_area                                  # :nodoc:
    attr_reader :writer, :set_rows, :col_info, :row_sizes         # :nodoc:
    attr_reader :vml_shape_id                                     # :nodoc:
    attr_reader :comments, :comments_author                       # :nodoc:
    attr_accessor :data_bars_2010, :dxf_priority                  # :nodoc:
    attr_reader :vba_codename                                     # :nodoc:
    attr_writer :excel_version                                    # :nodoc:
    attr_reader :filter_cells                                     # :nodoc:
    attr_accessor :default_row_height                             # :nodoc:

    def initialize(workbook, index, name) # :nodoc:
      @writer = Package::XMLWriterSimple.new

      setup_identity(workbook, index, name)
      setup_limits
      setup_dependencies
      setup_view_options
      setup_sheet_geometry
      setup_row_and_column_state
      setup_filter_and_selection_state
      setup_drawing_and_media
      setup_cell_features
      setup_protection

      apply_excel2003_compatibility if excel2003_style?
      setup_workbook_dependent_state
    end

    def set_xml_writer(filename) # :nodoc:
      @writer.set_xml_writer(filename)
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
    def hide(hidden = :hidden)
      @hidden = hidden
      @selected = false
      @workbook.activesheet = 0 if @workbook.activesheet == @index
      @workbook.firstsheet  = 0 if @workbook.firstsheet  == @index
    end

    #
    # Hide this worksheet. This can only be unhidden from VBA.
    #
    def very_hidden
      hide(:very_hidden)
    end

    def hidden? # :nodoc:
      @hidden == :hidden
    end

    def very_hidden? # :nodoc:
      @hidden == :very_hidden
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
        range = range.gsub("$", "")
        range = range.sub(/^=/, "")
        @num_protected_ranges += 1
      end

      range_name ||= "Range#{@num_protected_ranges}"
      password   &&= encode_password(password)

      @protected_ranges << [range, range_name, password]
    end

    #
    # The outline_settings() method is used to control the appearance of
    # outlines in Excel.
    #
    def outline_settings(visible = 1, symbols_below = 1, symbols_right = 1, auto_style = false)
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

    def_delegators :@assets,
                   :charts, :images, :tables, :shapes,
                   :header_images, :footer_images, :background_image

    #
    # :call-seq:
    #   insert_button(row, col, properties)
    #
    # The insert_button() method can be used to insert an Excel form button
    # into a worksheet.
    #
    def insert_button(row, col, properties = nil)
      if (row_col_array = row_col_notation(row))
        _row, _col = row_col_array
        _properties = col
      else
        _row = row
        _col = col
        _properties = properties
      end

      @buttons_array << Writexlsx::Package::Button.new(
        self, _row, _col, _properties, @default_row_pixels, @buttons_array.size + 1
      )
      @has_vml = true
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

    # autofilter methods moved to worksheet/autofilter.rb
    # see Writexlsx::Worksheet::Autofilter for implementation

    #
    # Store the horizontal page breaks on a worksheet.
    #

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
      !(header_images.empty? && footer_images.empty?)
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

    def comments_visible? # :nodoc:
      !!@comments_visible
    end

    def sorted_comments # :nodoc:
      @comments.sorted_comments
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
        "FF#{palette_color_from_index(index)}"
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
        @external_background_links,
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
      tables.size
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

    def has_dynamic_functions?
      @has_dynamic_functions
    end

    def has_embedded_images?
      @has_embedded_images
    end

    # Check that some image or drawing needs to be processed.
    def some_image_or_drawing_to_be_processed?
      charts.size + images.size + shapes.size + header_images.size + footer_images.size + (background_image ? 1 : 0) == 0
    end

    #
    # Set the background image for the worksheet.
    #
    def set_background(image)
      raise "Couldn't locate #{image}: $!" unless File.exist?(image)

      @assets.background_image = ImageProperty.new(image)
    end

    #
    # Calculate the vertices that define the position of a graphical object
    # within the worksheet in pixels.
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
                (0..(col_start - 1)).inject(0) { |sum, col| sum += size_col(col, anchor) }
              else
                # Optimisation for when the column widths haven't changed.
                DEFAULT_COL_PIXELS * col_start
              end
      x_abs += x1

      # Calculate the absolute y offset of the top-left vertex.
      # Store the column change to allow optimisations.
      y_abs = if @row_size_changed
                (0..(row_start - 1)).inject(0) { |sum, row| sum += size_row(row, anchor) }
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

    def date_1904? # :nodoc:
      @workbook.date_1904?
    end

    private

    def setup_identity(workbook, index, name)
      @workbook = workbook
      @index = index
      @name = name

      @excel_version = 2007
      @palette = workbook.palette
      @default_url_format = workbook.default_url_format
      @max_url_length = workbook.max_url_length
    end

    def setup_limits
      @xls_rowmax = 1_048_576
      @xls_colmax = 16_384
      @xls_strmax = 32_767

      @dim_rowmin = nil
      @dim_rowmax = nil
      @dim_colmin = nil
      @dim_colmax = nil
    end

    def setup_dependencies
      @page_setup = Writexlsx::PageSetup.new
      @comments   = Package::Comments.new(self)
      @assets     = AssetManager.new
    end

    def setup_view_options
      @screen_gridlines = true
      @show_zeros = true
      @hide_row_col_headers = 0
      @top_left_cell = ''

      @tab_color = 0

      @zoom = 100
      @zoom_scale_normal = true
      @right_to_left = false
      @leading_zeros = false
    end

    def setup_sheet_geometry
      @outline_row_level = 0
      @outline_col_level = 0

      @original_row_height = 15
      @default_row_height = 15
      @default_row_pixels = 20
      @default_col_width = 8.43
      @default_date_pixels = 68
    end

    def setup_row_and_column_state
      @col_info = {}
      @cell_data_store = CellDataStore.new

      @set_cols = {}
      @set_rows = {}
      @row_sizes = {}

      @col_size_changed = false
    end

    def setup_filter_and_selection_state
      @selections = []
      @panes = []

      @autofilter_area = nil
      @filter_on = false
      @filter_range = []
      @filter_cols = {}
      @filter_cells = {}
      @filter_type = {}
    end

    def setup_drawing_and_media
      @last_shape_id = 1
      @rel_count = 0

      @external_hyper_links = []
      @external_drawing_links = []
      @external_comment_links = []
      @external_vml_links = []
      @external_background_links = []
      @external_table_links = []

      @drawing_links = []
      @vml_drawing_links = []

      @shape_hash = {}
      @drawing_rels = {}
      @drawing_rels_id = 0
      @vml_drawing_rels = {}
      @vml_drawing_rels_id = 0

      @has_dynamic_functions = false
      @has_embedded_images = false
      @use_future_functions = false
      @has_vml = false

      @buttons_array = []
      @header_images_array = []
    end

    def setup_cell_features
      @merge = []

      @validations = []
      @cond_formats = {}
      @data_bars_2010 = []
      @dxf_priority = 1

      @ignore_errors = nil
    end

    def setup_protection
      @protected_ranges = []
      @num_protected_ranges = 0
    end

    def apply_excel2003_compatibility
      @original_row_height = 12.75
      @default_row_height = 12.75
      @default_row_pixels = 17

      self.margins_left_right = 0.75
      self.margins_top_bottom = 1

      @page_setup.margin_header = 0.5
      @page_setup.margin_footer = 0.5
      @page_setup.header_footer_aligns = false
    end

    def setup_workbook_dependent_state
      @embedded_image_indexes = @workbook.embedded_image_indexes
    end

    #
    # Convert the width of a cell from user's units to pixels. Excel rounds
    # the column width to the nearest pixel. If the width hasn't been set
    # by the user we use the default value. A hidden column is treated as
    # having a width of zero unless it has the special "object_position" of
    # 4 (size with cells).
    #
    def size_col(col, anchor = 0) # :nodoc:
      # Look up the cell value to see if it has been changed.
      if col_info[col]
        width  = col_info[col].width || @default_col_width
        hidden = col_info[col].hidden

        # Convert to pixels.
        pixels = if hidden == 1 && anchor != 4
                   0
                 elsif width < 1
                   ((width * (MAX_DIGIT_WIDTH + PADDING)) + 0.5).to_i
                 else
                   ((width * MAX_DIGIT_WIDTH) + 0.5).to_i + PADDING
                 end
      else
        pixels = DEFAULT_COL_PIXELS
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
      if row_sizes[row]
        height, hidden = row_sizes[row]

        pixels = if hidden == 1 && anchor != 4
                   0
                 else
                   (4 / 3.0 * height).to_i
                 end
      else
        pixels = (4 / 3.0 * default_row_height).to_i
      end
      pixels
    end

    #
    # Compare adjacent column information structures.
    #
    def compare_col_info(col_options, previous_options)
      if !col_options.width.nil? != !previous_options.width.nil?
        return nil
      end
      if col_options.width && previous_options.width &&
         col_options.width != previous_options.width
        return nil
      end

      if !col_options.format.nil? != !previous_options.format.nil?
        return nil
      end
      if col_options.format && previous_options.format &&
         col_options.format != previous_options.format
        return nil
      end

      return nil if col_options.hidden    != previous_options.hidden
      return nil if col_options.level     != previous_options.level
      return nil if col_options.collapsed != previous_options.collapsed

      true
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

      last = 'format'
      pos  = 0
      raw_string = ''

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

          raw_string += token    # Keep track of actual string length.
          last = 'string'
        end
        pos += 1
      end
      [fragments, raw_string]
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
        token.gsub!('""', '"')

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
    # Calculate the vertices that define the position of a graphical object
    # within the worksheet in EMUs.
    #
    def position_object_emus(graphical_object) # :nodoc:
      go = graphical_object
      col_start, row_start, x1, y1, col_end, row_end, x2, y2, x_abs, y_abs =
        position_object_pixels(go.col, go.row, go.x_offset, go.y_offset, go.scaled_width, go.scaled_height, go.anchor)

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
    # Hash a worksheet password. Based on the algorithm in ECMA-376-4:2016,
    # Office Open XML File Foemats -- Transitional Migration Features,
    # Additional attributes for workbookProtection element (Part 1, §18.2.29).   #
    def encode_password(password) # :nodoc:
      hash = 0

      password.reverse.split("").each do |char|
        hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7fff)
        hash ^= char.ord
      end

      hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7fff)
      hash ^= password.length
      hash ^= 0xCE4B

      sprintf("%X", hash)
    end

    def tab_outline_fit?
      tab_color? || outline_changed? || fit_page?
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
        span_min, span_max = calc_spans(@cell_data_store, row_num, span_min, span_max) if @cell_data_store[row_num]

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

    def protect_default_settings  # :nodoc:
      {
        sheet:                 true,
        content:               false,
        objects:               false,
        scenarios:             false,
        format_cells:          false,
        format_columns:        false,
        format_rows:           false,
        insert_columns:        false,
        insert_rows:           false,
        insert_hyperlinks:     false,
        delete_columns:        false,
        delete_rows:           false,
        select_locked_cells:   true,
        sort:                  false,
        autofilter:            false,
        pivot_tables:          false,
        select_unlocked_cells: true
      }
    end

    def expand_formula(formula, function, addition = '')
      if formula =~ /\b(#{function})/
        formula.gsub(
          ::Regexp.last_match(1),
          "_xlfn#{addition}.#{::Regexp.last_match(1)}"
        )
      else
        formula
      end
    end
  end
end
