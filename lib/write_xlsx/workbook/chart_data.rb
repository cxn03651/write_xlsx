# -*- coding: utf-8 -*-
# frozen_string_literal: true

module Writexlsx
  class Workbook
    module ChartData
      private

      #
      # Convert a range formula such as Sheet1!$B$1:$B$5 into a sheet name and cell
      # range such as ( 'Sheet1', 0, 1, 4, 1 ).
      #
      def get_chart_range(range) # :nodoc:
        # Split the range formula into sheetname and cells at the last '!'.
        pos = range.rindex('!')
        return nil unless pos

        if pos > 0
          sheetname = range[0, pos]
          cells = range[(pos + 1)..-1]
        end

        # Split the cell range into 2 cells or else use single cell for both.
        if cells =~ /:/
          cell_1, cell_2 = cells.split(":")
        else
          cell_1 = cells
          cell_2 = cells
        end

        # Remove leading/trailing apostrophes and convert escaped quotes to single.
        sheetname.sub!(/^'/, '')
        sheetname.sub!(/'$/, '')
        sheetname.gsub!("''", "'")

        row_start, col_start = xl_cell_to_rowcol(cell_1)
        row_end,   col_end   = xl_cell_to_rowcol(cell_2)

        # Check that we have a 1D range only.
        return nil if row_start != row_end && col_start != col_end

        [sheetname, row_start, col_start, row_end, col_end]
      end

      #
      # Add "cached" data to charts to provide the numCache and strCache data for
      # series and title/axis ranges.
      #
      def add_chart_data # :nodoc:
        worksheets = {}
        seen_ranges = {}

        # Map worksheet names to worksheet objects.
        @worksheets.each { |worksheet| worksheets[worksheet.name] = worksheet }

        # Build an array of the worksheet charts including any combined charts.
        @charts.collect { |chart| [chart, chart.combined] }.flatten.compact
          .each do |chart|
            chart.formula_ids.each do |range, id|
              # Skip if the series has user defined data.
              if chart.formula_data[id]
                seen_ranges[range] = chart.formula_data[id] unless seen_ranges[range]
                next
              # Check to see if the data is already cached locally.
              elsif seen_ranges[range]
                chart.formula_data[id] = seen_ranges[range]
                next
              end

              # Convert the range formula to a sheet name and cell range.
              sheetname, *cells = get_chart_range(range)

              # Skip if we couldn't parse the formula.
              next unless sheetname

              # Handle non-contiguous ranges: (Sheet1!$A$1:$A$2,Sheet1!$A$4:$A$5).
              # We don't try to parse the ranges. We just return an empty list.
              if sheetname =~ /^\([^,]+,/
                chart.formula_data[id] = []
                seen_ranges[range] = []
                next
              end

              # Raise if the name is unknown since it indicates a user error in
              # a chart series formula.
              raise "Unknown worksheet reference '#{sheetname}' in range '#{range}' passed to add_series()\n" unless worksheets[sheetname]

              # Add the data to the chart.
              # And store range data locally to avoid lookup if seen agein.
              chart.formula_data[id] =
                seen_ranges[range] = chart_data(worksheets[sheetname], cells)
            end
        end
      end

      def chart_data(worksheet, cells)
        # Get the data from the worksheet table.
        data = worksheet.get_range_data(*cells)

        # Convert shared string indexes to strings.
        data.collect do |token|
          if token.is_a?(Hash)
            string = @shared_strings.string(token[:sst_id])

            # Ignore rich strings for now. Deparse later if necessary.
            if string =~ /^<r>/ && string =~ %r{</r>$}
              ''
            else
              string
            end
          else
            token
          end
        end
      end

      #
      # Sort internal and user defined names in the same order as used by Excel.
      # This may not be strictly necessary but unsorted elements caused a lot of
      # issues in the the Spreadsheet::WriteExcel binary version. Also makes
      # comparison testing easier.
      #
      def sort_defined_names(names) # :nodoc:
        names.sort do |a, b|
          name_a  = normalise_defined_name(a[0])
          name_b  = normalise_defined_name(b[0])
          sheet_a = normalise_sheet_name(a[2])
          sheet_b = normalise_sheet_name(b[2])
          # Primary sort based on the defined name.
          if name_a > name_b
            1
          elsif name_a < name_b
            -1
          elsif sheet_a >= sheet_b  # name_a == name_b
            # Secondary sort based on the sheet name.
            1
          else
            -1
          end
        end
      end

      # Used in the above sort routine to normalise the defined names. Removes any
      # leading '_xmln.' from internal names and lowercases the strings.
      def normalise_defined_name(name) # :nodoc:
        name.sub(/^_xlnm./, '').downcase
      end

      # Used in the above sort routine to normalise the worksheet names for the
      # secondary sort. Removes leading quote and lowercases the strings.
      def normalise_sheet_name(name) # :nodoc:
        name.sub(/^'/, '').downcase
      end

      #
      # Extract the named ranges from the sorted list of defined names. These are
      # used in the App.xml file.
      #
      def extract_named_ranges(defined_names) # :nodoc:
        named_ranges = []

        defined_names.each do |defined_name|
          name, index, range = defined_name

          # Skip autoFilter ranges.
          next if name == '_xlnm._FilterDatabase'

          # We are only interested in defined names with ranges.
          next unless range =~ /^([^!]+)!/

          sheet_name = ::Regexp.last_match(1)

          # Match Print_Area and Print_Titles xlnm types.
          if name =~ /^_xlnm\.(.*)$/
            xlnm_type = ::Regexp.last_match(1)
            name = "#{sheet_name}!#{xlnm_type}"
          elsif index != -1
            name = "#{sheet_name}!#{name}"
          end

          named_ranges << name
        end

        named_ranges
      end
    end
  end
end
