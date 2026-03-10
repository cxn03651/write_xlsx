# frozen_string_literal: true

module Writexlsx
  class Worksheet
    module XmlWriter
      include Utility

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
      # Write the cell value <v> element.
      #
      def write_cell_value(value = '') # :nodoc:
        return write_cell_formula('=NA()') if value.is_a?(Float) && value.nan?

        value ||= ''

        int_value = value.to_i
        value = int_value if value == int_value
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
        case @page_view
        when 1
          attributes << %w[view pageLayout]
        when 2
          attributes << %w[view pageBreakPreview]
        end

        # Set the first visible cell.
        attributes << ['topLeftCell', @top_left_cell] if ptrue?(@top_left_cell)

        # Set the zoom level.
        if @zoom != 100
          attributes << ['zoomScale', @zoom]

          if @page_view == 1
            attributes << ['zoomScalePageLayoutView', @zoom]
          elsif @page_view == 2
            attributes << ['zoomScaleSheetLayoutView', @zoom]
          elsif ptrue?(@zoom_scale_normal)
            attributes << ['zoomScaleNormal', @zoom]
          end
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
      # Write the <cols> element and <col> sub elements.
      #
      def write_cols # :nodoc:
        # Exit unless some column have been formatted.
        return if @col_info.empty?

        @writer.tag_elements('cols') do
          # Use the first element of the column informatin structure to set
          # the initial/previous properties.
          first_col           = @col_info.keys.min
          last_col            = first_col
          previous_options    = @col_info[first_col]
          deleted_col         = first_col
          deleted_col_options = previous_options

          @col_info.delete(first_col)

          @col_info.keys.sort.each do |col|
            col_options = @col_info[col]

            # Check if the column number is contiguous with the previous
            # column and if the properties are the same.
            if (col == last_col + 1) &&
               compare_col_info(col_options, previous_options)
              last_col = col
            else
              # If not contiguous/equal then we write out the current range
              # of columns and start again.
              write_col_info([first_col, last_col, previous_options])
              first_col = col
              last_col  = first_col
              previous_options = col_options
            end
          end

          # We will exit the previous loop with one unhandled column range.
          write_col_info([first_col, last_col, previous_options])

          # Put back the deleted first column information structure:
          @col_info[deleted_col] = deleted_col_options
        end
      end

      #
      # Write the <col> element.
      #
      def write_col_info(args) # :nodoc:
        @writer.empty_tag('col', col_info_attributes(args))
      end

      def col_info_attributes(args)
        min       = args[0]           || 0 # First formatted column.
        max       = args[1]           || 0 # Last formatted column.
        width     = args[2].width          # Col width in user units.
        format    = args[2].format         # Format index.
        hidden    = args[2].hidden    || 0 # Hidden flag.
        level     = args[2].level     || 0 # Outline level.
        collapsed = args[2].collapsed || 0 # Outline Collapsed
        autofit   = args[2].autofit   || 0 # Best fit for autofit numbers.
        xf_index = format ? format.get_xf_index : 0

        custom_width = true
        custom_width = false if width.nil? && hidden == 0
        custom_width = false if !width.nil? && (width - 8.43).abs < 0.01

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

        attributes << ['style',        xf_index] if xf_index  != 0
        attributes << ['hidden',       1]        if hidden    != 0
        attributes << ['bestFit',      1]        if autofit   != 0
        attributes << ['customWidth',  1]        if custom_width
        attributes << ['outlineLevel', level]    if level     != 0
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
          if @cell_data_store[row_num]
            args = @set_rows[row_num] || []
            write_row_element(row_num, span, *args) do
              write_cell_column_dimension(row_num)
            end
          else
            # Row attributes only.
            write_empty_row(row_num, span, *@set_rows[row_num])
          end
        end
      end

      def not_contain_formatting_or_data?(row_num) # :nodoc:
        !@set_rows[row_num] && !@cell_data_store[row_num] && !@comments.has_comment_in_row?(row_num)
      end

      #
      # Write the <row> element.
      #
      def write_row_element(*args, &block)  # :nodoc:
        @writer.tag_elements('row', row_attributes(args), &block)
      end

      def write_cell_column_dimension(row_num)  # :nodoc:
        row = @cell_data_store[row_num]
        row_name = (row_num + 1).to_s
        (@dim_colmin..@dim_colmax).each do |col_num|
          if (cell = row[col_num])
            cell.write_cell(self, row_num, row_name, col_num)
          end
        end
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
      # Write the Worksheet conditional formats.
      #
      def write_conditional_formats  # :nodoc:
        @cond_formats.keys.sort.each do |range|
          write_conditional_formatting(range, @cond_formats[range])
        end
      end

      # conditional formatting XML writing moved to worksheet/conditional_formats.rb
      # see Writexlsx::Worksheet::ConditionalFormats#write_conditional_formatting

      #
      # Write the <dataValidations> element.
      #
      def write_data_validations # :nodoc:
        write_some_elements('dataValidations', @validations) do
          @validations.each { |validation| validation.write_data_validation(@writer) }
        end
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
            if ptrue?(@cell_data_store)                   &&
               ptrue?(@cell_data_store[row_num])          &&
               ptrue?(@cell_data_store[row_num][col_num]) &&
               @cell_data_store[row_num][col_num].display_url_string?
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
      # Write the <printOptions> element.
      #
      def write_print_options # :nodoc:
        @page_setup.write_print_options(@writer)
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
      # Write the <ignoredErrors> element.
      #
      def write_ignored_errors
        return unless @ignore_errors

        ignore = @ignore_errors

        @writer.tag_elements('ignoredErrors') do
          {
            number_stored_as_text: 'numberStoredAsText',
            eval_error:            'evalError',
            formula_differs:       'formula',
            formula_range:         'formulaRange',
            formula_unlocked:      'unlockedFormula',
            empty_cell_reference:  'emptyCellReference',
            list_data_validation:  'listDataValidation',
            calculated_column:     'calculatedColumn',
            two_digit_text_year:   'twoDigitTextYear'
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
        return unless @background_image

        # Increment the relationship id.
        @rel_count += 1
        id = @rel_count

        attributes = [['r:id', "rId#{id}"]]

        @writer.empty_tag('picture', attributes)
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
          write_ext_list_data_bars  unless @data_bars_2010.empty?
          write_ext_list_sparklines unless @sparklines.empty?
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

      def write_ext(url, &block)
        attributes = [
          ['xmlns:x14', "#{OFFICE_URL}spreadsheetml/2009/9/main"],
          ['uri',       url]
        ]
        @writer.tag_elements('ext', attributes, &block)
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
        attributes << ["applyStyles",  1] if @outline_style
        attributes << ["summaryBelow", 0] if @outline_below == 0
        attributes << ["summaryRight", 0] if @outline_right == 0
        attributes << ["showOutlineSymbols", 0] if @outline_on == 0

        @writer.empty_tag('outlinePr', attributes)
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
    end
  end
end
