# -*- coding: utf-8 -*-
require 'write_xlsx/package/xml_writer_simple'
require 'write_xlsx/package/packager'
require 'write_xlsx/worksheet'
require 'write_xlsx/format'
require 'write_xlsx/utility'
require 'write_xlsx/chart'
require 'write_xlsx/zip_file_utils'
require 'tmpdir'
require 'tempfile'
require 'digest/md5'

module Writexlsx
  class Workbook

    include Utility

    attr_accessor :str_total, :str_unique, :str_table
    attr_writer :firstsheet
    attr_reader :palette
    attr_reader :font_count, :num_format_count, :border_count, :fill_count, :custom_colors
    attr_reader :xf_formats, :dxf_formats
    attr_reader :worksheets, :sheetnames, :charts, :drawings, :num_comment_files, :named_ranges
    attr_reader :str_array, :doc_properties
    attr_reader :image_types, :images
    attr_reader :named_ranges

    def initialize(file, default_formats = {})
      @writer = Package::XMLWriterSimple.new

      @tempdir  = File.join(Dir.tmpdir, Digest::MD5.hexdigest(Time.now.to_s))
      setup_filename(file)
      @date_1904           = false
      @activesheet         = 0
      @firstsheet          = 0
      @selected            = 0
      @fileclosed          = false
      @sheet_name          = 'Sheet'
      @chart_name          = 'Chart'
      @sheetname_count     = 0
      @chartname_count     = 0
      @worksheets          = []
      @charts              = []
      @drawings            = []
      @sheetnames          = []
      @formats             = []
      @xf_formats          = []
      @xf_format_indices   = {}
      @dxf_formats         = []
      @dxf_format_indices  = {}
      @font_count          = 0
      @num_format_count    = 0
      @defined_names       = []
      @named_ranges        = []
      @custom_colors       = []
      @doc_properties      = {}
      @local_time          = Time.now
      @num_comment_files   = 0
      @image_types         = {}
      @images              = []

      # Structures for the shared strings data.
      @str_total  = 0
      @str_unique = 0
      @str_table  = {}
      @str_array  = []

      add_format(default_formats.merge(:xf_index => 0))
      set_color_palette
    end

    #
    # Calls finalization methods.
    #
    def close
      # In case close() is called twice, by user and by DESTROY.
      return if @fileclosed

      # Test filehandle in case new() failed and the user didn't check.
#          return nil unless @filehandle

      @fileclosed = 1
      store_workbook
    end

    def set_xml_writer(filename)
      @writer.set_xml_writer(filename)
    end

    def xml_str
      @writer.string
    end

    def assemble_xml_file
      return unless @writer

      # Prepare format object for passing to Style.pm.
      prepare_format_properties

      write_xml_declaration

      # Write the root workbook element.
      write_workbook

      # Write the XLSX file version.
      write_file_version

      # Write the workbook properties.
      write_workbook_pr

      # Write the workbook view properties.
      write_book_views

      # Write the worksheet names and ids.
      write_sheets

      # Write the workbook defined names.
      write_defined_names

      # Write the workbook calculation properties.
      write_calc_pr

      # Write the workbook extension storage.
      #write_ext_lst

      # Close the workbook tag.
      write_workbook_end

      # Close the XML writer object and filehandle.
      @writer.crlf
      @writer.close
    end

    def add_worksheet(name = '')
      index = @worksheets.size
      name  = check_sheetname(name)

      worksheet = Worksheet.new(self, index, name)
      @worksheets[index] = worksheet
      @sheetnames[index] = name

      worksheet
    end

    #
    # Create a chart for embedding or as as new sheet.
    #
    def add_chart(params)
      name     = ''
      index    = @worksheets.size

      # Type must be specified so we can create the required chart instance.
      type = params[:type]
      raise "Must define chart type in add_chart()" unless type

      # Ensure that the chart defaults to non embedded.
      embedded = params[:embedded] || 0

      # Check the worksheet name for non-embedded charts.
      name = check_sheetname(params[:name], 1) unless embedded

      chart = Chart.factory(type)

      # Get an incremental id to use for axes ids.
      chart.id = @charts.size

      # If the chart isn't embedded let the workbook control it.
      if embedded
        # Set index to 0 so that the activate() and set_first_sheet() methods
        # point back to the first worksheet if used for embedded charts.
        chart.index = 0
        chart.palette = @palette
        chart.set_embedded_config_data
        @charts << chart

        return chart
      else
        chartsheet = Chartsheet.new(self, name, index)
        chart.palette = @palette
        chartsheet.chart   = chart
        chartsheet.drawing = Drawing.new
        @worksheets.index = chartsheet
        @sheetnames.index = name

        @charts << chart

        return chartsheet
      end
    end

    #
    # Add a new format to the Excel workbook.
    #
    def add_format(properties)
      init_data = [
        @xf_format_indices,
        @dxf_format_indices,
        properties
      ]

      format = Format.new(*init_data)

      @formats.push(format)    # Store format reference

      format
    end

    #
    # Sets the colour palette to the Excel defaults.
    #
    def set_color_palette
      @palette = [
            [ 0x00, 0x00, 0x00, 0x00 ],    # 8
            [ 0xff, 0xff, 0xff, 0x00 ],    # 9
            [ 0xff, 0x00, 0x00, 0x00 ],    # 10
            [ 0x00, 0xff, 0x00, 0x00 ],    # 11
            [ 0x00, 0x00, 0xff, 0x00 ],    # 12
            [ 0xff, 0xff, 0x00, 0x00 ],    # 13
            [ 0xff, 0x00, 0xff, 0x00 ],    # 14
            [ 0x00, 0xff, 0xff, 0x00 ],    # 15
            [ 0x80, 0x00, 0x00, 0x00 ],    # 16
            [ 0x00, 0x80, 0x00, 0x00 ],    # 17
            [ 0x00, 0x00, 0x80, 0x00 ],    # 18
            [ 0x80, 0x80, 0x00, 0x00 ],    # 19
            [ 0x80, 0x00, 0x80, 0x00 ],    # 20
            [ 0x00, 0x80, 0x80, 0x00 ],    # 21
            [ 0xc0, 0xc0, 0xc0, 0x00 ],    # 22
            [ 0x80, 0x80, 0x80, 0x00 ],    # 23
            [ 0x99, 0x99, 0xff, 0x00 ],    # 24
            [ 0x99, 0x33, 0x66, 0x00 ],    # 25
            [ 0xff, 0xff, 0xcc, 0x00 ],    # 26
            [ 0xcc, 0xff, 0xff, 0x00 ],    # 27
            [ 0x66, 0x00, 0x66, 0x00 ],    # 28
            [ 0xff, 0x80, 0x80, 0x00 ],    # 29
            [ 0x00, 0x66, 0xcc, 0x00 ],    # 30
            [ 0xcc, 0xcc, 0xff, 0x00 ],    # 31
            [ 0x00, 0x00, 0x80, 0x00 ],    # 32
            [ 0xff, 0x00, 0xff, 0x00 ],    # 33
            [ 0xff, 0xff, 0x00, 0x00 ],    # 34
            [ 0x00, 0xff, 0xff, 0x00 ],    # 35
            [ 0x80, 0x00, 0x80, 0x00 ],    # 36
            [ 0x80, 0x00, 0x00, 0x00 ],    # 37
            [ 0x00, 0x80, 0x80, 0x00 ],    # 38
            [ 0x00, 0x00, 0xff, 0x00 ],    # 39
            [ 0x00, 0xcc, 0xff, 0x00 ],    # 40
            [ 0xcc, 0xff, 0xff, 0x00 ],    # 41
            [ 0xcc, 0xff, 0xcc, 0x00 ],    # 42
            [ 0xff, 0xff, 0x99, 0x00 ],    # 43
            [ 0x99, 0xcc, 0xff, 0x00 ],    # 44
            [ 0xff, 0x99, 0xcc, 0x00 ],    # 45
            [ 0xcc, 0x99, 0xff, 0x00 ],    # 46
            [ 0xff, 0xcc, 0x99, 0x00 ],    # 47
            [ 0x33, 0x66, 0xff, 0x00 ],    # 48
            [ 0x33, 0xcc, 0xcc, 0x00 ],    # 49
            [ 0x99, 0xcc, 0x00, 0x00 ],    # 50
            [ 0xff, 0xcc, 0x00, 0x00 ],    # 51
            [ 0xff, 0x99, 0x00, 0x00 ],    # 52
            [ 0xff, 0x66, 0x00, 0x00 ],    # 53
            [ 0x66, 0x66, 0x99, 0x00 ],    # 54
            [ 0x96, 0x96, 0x96, 0x00 ],    # 55
            [ 0x00, 0x33, 0x66, 0x00 ],    # 56
            [ 0x33, 0x99, 0x66, 0x00 ],    # 57
            [ 0x00, 0x33, 0x00, 0x00 ],    # 58
            [ 0x33, 0x33, 0x00, 0x00 ],    # 59
            [ 0x99, 0x33, 0x00, 0x00 ],    # 60
            [ 0x99, 0x33, 0x66, 0x00 ],    # 61
            [ 0x33, 0x33, 0x99, 0x00 ],    # 62
            [ 0x33, 0x33, 0x33, 0x00 ],    # 63
        ]
    end

    #
    # Create a defined name in Excel. We handle global/workbook level names and
    # local/worksheet names.
    #
    def define_name(name, formula)
      sheet_index = nil
      sheetname   = ''
      full_name   = name

      # Remove the = sign from the formula if it exists.
      formula.sub!(/^=/, '')

      # Local defined names are formatted like "Sheet1!name".
      if name =~ /^(.*)!(.*)$/
        sheetname   = $1
        name        = $2
        sheet_index = get_sheet_index(sheetname)
      else
        sheet_index =-1   # Use -1 to indicate global names.
      end

      # Warn if the sheet index wasn't found.
      if !sheet_index
       raise "Unknown sheet name #{sheetname} in defined_name()\n"
       return -1
      end

      # Warn if the sheet name contains invalid chars as defined by Excel help.
      if name !~ %r!^[a-zA-Z_\\][a-zA-Z_.]+!
       raise "Invalid characters in name '#{name}' used in defined_name()\n"
       return -1
      end

      # Warn if the sheet name looks like a cell name.
      if name =~ %r(^[a-zA-Z][a-zA-Z]?[a-dA-D]?[0-9]+$)
        raise "Invalid name '#{name}' looks like a cell name in defined_name()\n"
        return -1
      end

      @defined_names.push([ name, sheet_index, formula])
    end

    #
    # Set the document properties such as Title, Author etc. These are written to
    # property sets in the OLE container.
    #
    def set_properties(params)
      # Ignore if no args were passed.
      return -1 if params.empty?

      # List of valid input parameters.
      valid = {
        :title       => 1,
        :subject     => 1,
        :author      => 1,
        :keywords    => 1,
        :comments    => 1,
        :last_author => 1,
        :created     => 1,
        :category    => 1,
        :manager     => 1,
        :company     => 1,
        :status      => 1
      }

      # Check for valid input parameters.
      params.each_key do |key|
        return -1 unless valid.has_key?(key)
      end

      # Set the creation time unless specified by the user.
      params[:created] = @local_time unless params.has_key?(:created)

      @doc_properties = params.dup
    end

    def activesheet=(worksheet)
      @activesheet = worksheet
    end

    def writer
      @writer
    end

    attr_writer :date_1904

    def date_1904?
      @date_1904 ||= false
      !!@date_1904
    end

    private

    def setup_filename(file)
      if file.respond_to?(:to_str) && file != ''
        @filename = file
        @fileobj  = nil
      elsif file.respond_to?(:write)
        @filename = File.join(@tempdir, Digest::MD5.hexdigest(Time.now.to_s) + '.xlsx.tmp')
        @fileobj  = file
      else
        raise "'file' must be valid filename String of IO object."
      end
    end

    #
    # Check for valid worksheet names. We check the length, if it contains any
    # invalid characters and if the name is unique in the workbook.
    #
    def check_sheetname(name, chart = nil)
      name  ||= ''
      invalid_char = /[\[\]:*?\/\\]/

      # Increment the Sheet/Chart number used for default sheet names below.
      if chart
        @chartname_count += 1
      else
        @sheetname_count += 1
      end

      # Supply default Sheet/Chart name if none has been defined.
      if name == ''
        if chart
          name = "#{@chart_name}#{@chartname_count}"
        else
          name = "#{@sheet_name}#{@sheetname_count}"
        end
      end

      # Check that sheet name is <= 31. Excel limit.
      raise "Sheetname #{name} must be <= 31 chars" if name.bytesize > 31

      # Check that sheetname doesn't contain any invalid characters
      if name =~ invalid_char
        raise 'Invalid character []:*?/\\ in worksheet name: ' + name
      end

      # Check that the worksheet name doesn't already exist since this is a fatal
      # error in Excel 97. The check must also exclude case insensitive matches.
      @worksheets.each do |worksheet|
        name_a = name
        name_b = worksheet.name

        if name_a.downcase == name_b.downcase
          raise "Worksheet name '#{name}', with case ignored, is already used."
        end
      end

      name
    end

    #
    # Convert a range formula such as Sheet1!$B$1:$B$5 into a sheet name and cell
    # range such as ( 'Sheet1', 0, 1, 4, 1 ).
    #
    def get_chart_range(range)
      # Split the range formula into sheetname and cells at the last '!'.
      pos = range.rindex('!')
      return nil unless pos

      if pos > 0
        sheetname = range[0, pos]
        cells = range[pos + 1 .. -1]
      end

      # Split the cell range into 2 cells or else use single cell for both.
      if cells =~ /:/
        cell_1, cell_2 = cells.split(/:/)
      else
        cell_1, cell_2 = cells, cells
      end

      # Remove leading/trailing apostrophes and convert escaped quotes to single.
      sheetname.sub!(/^'/, '')
      sheetname.sub!(/'$/, '')
      sheetname.gsub!(/''/, "'")

      row_start, col_start = xl_cell_to_rowcol(cell_1)
      row_end,   col_end   = xl_cell_to_rowcol(cell_2)

      # Check that we have a 1D range only.
      return nil if row_start != row_end && col_start != col_end
      return [sheetname, row_start, col_start, row_end, col_end]
    end

    def write_xml_declaration
      @writer.xml_decl
    end

    def write_workbook
      schema  = 'http://schemas.openxmlformats.org'
      attributes = [
        'xmlns',
        schema + '/spreadsheetml/2006/main',
        'xmlns:r',
        schema + '/officeDocument/2006/relationships'
      ]
      @writer.start_tag('workbook', attributes)
    end

    def write_workbook_end
      @writer.end_tag('workbook')
    end

    def write_file_version
      attributes = [
        'appName', 'xl',
        'lastEdited', 4,
        'lowestEdited', 4,
        'rupBuild', 4505
      ]
      @writer.empty_tag('fileVersion', attributes)
    end

    def write_workbook_pr
      attributes = date_1904? ? ['date1904', 1] : []
      attributes << 'defaultThemeVersion' << 124226
      @writer.empty_tag('workbookPr', attributes)
    end

    def write_book_views
      @writer.start_tag('bookViews') << write_workbook_view << @writer.end_tag('bookViews')
    end

    def write_workbook_view
      attributes = [
        'xWindow',        240,
        'yWindow',         15,
        'windowWidth',  16095,
        'windowHeight',  9660
      ]
      if @firstsheet > 0
        attributes << 'firstSheet' << @firstsheet
      end
      if @activesheet > 0
        attributes << 'activeTab' << @activesheet
      end
      @writer.empty_tag('workbookView', attributes)
    end

    def write_sheets
      str = @writer.start_tag('sheets')
      id_num = 1
      @worksheets.each do |sheet|
        str << write_sheet(sheet.name, id_num, sheet.hidden)
        id_num += 1
      end
      str << @writer.end_tag('sheets')
    end

    def write_sheet(name, sheet_id, hidden = false)
      attributes = [
        'name',    name,
        'sheetId', sheet_id
      ]

      if hidden
        attributes << 'state' << 'hidden'
      end
      attributes << 'r:id' << "rId#{sheet_id}"
      @writer.empty_tag('sheet', attributes)
    end

    def write_calc_pr
      attributes = ['calcId', 124519]
      @writer.empty_tag('calcPr', attributes)
    end

    def write_ext_lst
      tag = 'extLst'
      @writer.start_tag(tag) << write_ext << @writer.end_tag(tag)
    end

    def write_ext
      tag = 'ext'
      attributes = [
        'xmlns:mx', 'http://schemas.microsoft.com/office/mac/excel/2008/main',
        'uri', 'http://schemas.microsoft.com/office/mac/excel/2008/main'
      ]
      @writer.start_tag(tag, attributes) << write_mx_arch_id << @writer.end_tag(tag)
    end

    def write_mx_arch_id
      @writer.empty_tag('mx:ArchID', ['Flags', 2])
    end

    def write_defined_names
      return if @defined_names.nil? || @defined_names.empty?
      tag = 'definedNames'
      str = @writer.start_tag(tag)
      @defined_names.each { |defined_name| str << write_defined_name(defined_name) }
      str << @writer.end_tag(tag)
    end

    def write_defined_name(data)
      name, id, range, hidden = data

      attributes = ['name', name]
      attributes << 'localSheetId' << "#{id}" unless id == -1
      attributes << 'hidden'       << '1'     if hidden

      @writer.data_element('definedName', range, attributes)
    end

    def write_io(str)
      @writer << str
      str
    end

    def firstsheet
      @firstsheet ||= 0
    end

    def activesheet
      @activesheet ||= 0
    end

    # for test
    def defined_names
      @defined_names ||= []
    end

    #
    # Assemble worksheets into a workbook.
    #
    def store_workbook
      packager = Package::Packager.new

      # Add a default worksheet if non have been added.
      add_worksheet if @worksheets.empty?

      # Ensure that at least one worksheet has been selected.
      @worksheets.first.select if @activesheet == 0

      # Set the active sheet.
      @worksheets.each { |sheet| sheet.activate if sheet.index == @activesheet }

      # Convert the SST strings data structure.
      prepare_sst_string_data

      # Prepare the worksheet cell comments.
      prepare_comments

      # Set the defined names for the worksheets such as Print Titles.
      prepare_defined_names

      # Prepare the drawings, charts and images.
      prepare_drawings

      # Add cached data to charts.
      add_chart_data

      # Package the workbook.
      packager.add_workbook(self)
      packager.set_package_dir(@tempdir)
      packager.create_package

      # Free up the Packager object.
      packager = nil

      # Store the xlsx component files with the temp dir name removed.
      ZipFileUtils.zip("#{@tempdir}", @filename)
      IO.copy_stream(@filename, @fileobj) if @fileobj
      Utility.delete_files(@tempdir)
    end

    #
    # Convert the SST string data from a hash to an array.
    #
    def prepare_sst_string_data
      strings = []

      @str_table.each_key { |key| strings[@str_table[key]] = key }

      # The SST data could be very large, free some memory (maybe).
      @str_table = nil
      @str_array = strings
    end

    #
    # Prepare all of the format properties prior to passing them to Styles.pm.
    #
    def prepare_format_properties
      # Separate format objects into XF and DXF formats.
      prepare_formats

      # Set the font index for the format objects.
      prepare_fonts

      # Set the number format index for the format objects.
      prepare_num_formats

      # Set the border index for the format objects.
      prepare_borders

      # Set the fill index for the format objects.
      prepare_fills
    end

    #
    # Iterate through the XF Format objects and separate them into XF and DXF
    # formats.
    #
    def prepare_formats
      @formats.each do |format|
        xf_index  = format.xf_index
        dxf_index = format.dxf_index

        @xf_formats[xf_index] = format   if xf_index
        @dxf_formats[dxf_index] = format if dxf_index
      end
    end

    #
    # Set the default index for each format. This is mainly used for testing.
    #
    def set_default_xf_indices
      @formats.each { |format| format.get_xf_index }
    end

    #
    # Iterate through the XF Format objects and give them an index to non-default
    # font elements.
    #
    def prepare_fonts
      fonts = {}
      index = 0

      @xf_formats.each do |format|
        key = format.get_font_key

        if fonts[key]
          # Font has already been used.
          format.font_index = fonts[key]
          format.has_font   = 0
        else
          # This is a new font.
          fonts[key]        = index
          format.font_index = index
          format.has_font   = 1
          index += 1
        end
      end

      @font_count = index

      # For the DXF formats we only need to check if the properties have changed.
      @dxf_formats.each do |format|
        # The only font properties that can change for a DXF format are: color,
        # bold, italic, underline and strikethrough.
        if format.color || format.bold || format.italic || format.underline || format.font_strikeout
          format.has_dxf_font = 1
        end
      end
    end

    #
    # Iterate through the XF Format objects and give them an index to non-default
    # number format elements.
    #
    # User defined records start from index 0xA4.
    #
    def prepare_num_formats
      num_formats      = {}
      index            = 164
      num_format_count = 0

      (@xf_formats + @dxf_formats).each do |format|
        num_format = format.num_format

        # Check if num_format is an index to a built-in number format.
        # Also check for a string of zeros, which is a valid number format
        # string but would evaluate to zero.
        #
        if num_format.to_s =~ /^\d+$/ && num_format.to_s !~ /^0+\d/
          # Index to a built-in number format.
          format.num_format_index = num_format
          next
        end

        if num_formats[num_format]
          # Number format has already been used.
          format.num_format_index = num_formats[num_format]
        else
          # Add a new number format.
          num_formats[num_format] = index
          format.num_format_index = index
          index += 1

          # Only increase font count for XF formats (not for DXF formats).
          num_format_count += 1 unless format.xf_index == 0
        end
      end

      @num_format_count = num_format_count
    end

    #
    # Iterate through the XF Format objects and give them an index to non-default
    # border elements.
    #
    def prepare_borders
      borders = {}
      index = 0

      @xf_formats.each do |format|
        key = format.get_border_key

        if borders[key]
          # Border has already been used.
          format.border_index = borders[key]
          format.has_border   = 0
        else
          # This is a new border.
          borders[key]        = index
          format.border_index = index
          format.has_border   = 1
          index += 1
        end
      end

      @border_count = index

      # For the DXF formats we only need to check if the properties have changed.
      @dxf_formats.each do |format|
        key = format.get_border_key
        format.has_dxf_border = 1 if key =~ /[^0:]/
      end
    end

    #
    # Iterate through the XF Format objects and give them an index to non-default
    # fill elements.
    #
    # The user defined fill properties start from 2 since there are 2 default
    # fills: patternType="none" and patternType="gray125".
    #
    def prepare_fills
      fills = {}
      index = 2    # Start from 2. See above.

      # Add the default fills.
      fills['0:0:0']  = 0
      fills['17:0:0'] = 1

      @xf_formats.each do |format|
        # The following logical statements jointly take care of special cases
        # in relation to cell colours and patterns:
        # 1. For a solid fill (_pattern == 1) Excel reverses the role of
        #    foreground and background colours, and
        # 2. If the user specifies a foreground or background colour without
        #    a pattern they probably wanted a solid fill, so we fill in the
        #    defaults.
        #
        if format.pattern <= 1 && format.bg_color != 0 && format.fg_color == 0
          format.fg_color = format.bg_color
          format.bg_color = 0
          format.pattern  = 1
        end

        if format.pattern <= 1 && format.bg_color == 0 && format.fg_color != 0
          format.bg_color = 0
          format.pattern  = 1
        end

        key = format.get_fill_key

        if fills[key]
          # Fill has already been used.
          format.fill_index = fills[key]
          format.has_fill   = 0
        else
          # This is a new fill.
          fills[key]        = index
          format.fill_index = index
          format.has_fill   = 1
          index += 1
        end
      end

      @fill_count = index

      # For the DXF formats we only need to check if the properties have changed.
      @dxf_formats.each do |format|
        format.has_dxf_fill = 1 if format.pattern || format.bg_color || format.fg_color
      end
    end

    #
    # Iterate through the worksheets and store any defined names in addition to
    # any user defined names. Stores the defined names for the Workbook.xml and
    # the named ranges for App.xml.
    #
    def prepare_defined_names
      defined_names =  @defined_names

      @worksheets.each do |sheet|
        # Check for Print Area settings.
        if sheet.autofilter_area
          range  = sheet.autofilter_area
          hidden = 1

          # Store the defined names.
          defined_names << ['_xlnm._FilterDatabase', sheet.index, range, hidden]
        end

        # Check for Print Area settings.
        if !sheet.print_area.empty?
          range = sheet.print_area

          # Store the defined names.
          defined_names << ['_xlnm.Print_Area', sheet.index, range]
        end

        # Check for repeat rows/cols. aka, Print Titles.
        if !sheet._repeat_cols.empty? || !sheet._repeat_rows.empty?
          range = ''

          if !sheet._repeat_cols.empty? && !sheet._repeat_rows.empty?
            range = sheet._repeat_cols + ',' + sheet._repeat_rows
          else
            range = sheet._repeat_cols + sheet._repeat_rows
          end

          # Store the defined names.
          defined_names << ['_xlnm.Print_Titles', sheet.index, range]
        end
      end

      defined_names  = sort_defined_names(defined_names)
      @defined_names = defined_names
      @named_ranges  = extract_named_ranges(defined_names)
    end

    #
    # Iterate through the worksheets and set up the comment data.
    #
    def prepare_comments
      comment_id   = 0
      vml_data_id  = 1
      vml_shape_id = 1024

      @worksheets.each do |sheet|
        next unless sheet.has_comments?

        comment_id += 1
        count = sheet.prepare_comments( vml_data_id, vml_shape_id, comment_id)

        # Each VML file should start with a shape id incremented by 1024.
        vml_data_id  +=    1 * ( ( 1024 + count ) / 1024.0 ).to_i
        vml_shape_id += 1024 * ( ( 1024 + count ) / 1024.0 ).to_i
      end

      @num_comment_files = comment_id

      # Add a font format for cell comments.
      if comment_id > 0
        format = Format.new(
            @xf_format_indices,
            @dxf_format_indices,
            :font          => 'Tahoma',
            :size          => 8,
            :color_indexed => 81,
            :font_only     => 1
        )

        format.get_xf_index

        @formats << format
      end
    end

    #
    # Add "cached" data to charts to provide the numCache and strCache data for
    # series and title/axis ranges.
    #
    def add_chart_data
      worksheets = {}
      seen_ranges = {}

      # Map worksheet names to worksheet objects.
      @worksheets.each { |worksheet| worksheets[worksheet.name] = worksheet }

      @charts.each do |chart|
        chart.formula_ids.each do |range, id|
          # Skip if the series has user defined data.
          if chart.formula_data[id]
            if !seen_ranges.has_key?(range) || seen_ranges[range]
              data = chart.formula_data[id]
              seen_ranges[range] = data
            end
            next
          end

          # Check to see if the data is already cached locally.
          if seen_ranges.has_key?(range)
            chart.formula_data[id] = seen_ranges[range]
            next
          end

          # Convert the range formula to a sheet name and cell range.
          sheetname, *cells = get_chart_range(range)

          # Skip if we couldn't parse the formula.
          next unless sheetname

          # Skip if the name is unknown. Probably should throw exception.
          next unless worksheets[sheetname]

          # Find the worksheet object based on the sheet name.
          worksheet = worksheets[sheetname]

          # Get the data from the worksheet table.
          data = worksheet.get_range_data(*cells)

          # Convert shared string indexes to strings.
          data.collect! do |token|
            if token.kind_of?(Hash)
              token = @str_array[token[:sst_id]]

              # Ignore rich strings for now. Deparse later if necessary.
              token = '' if token =~ %r!^<r>! && token =~ %r!</r>$!
            end
            token
          end

          # Add the data to the chart.
          chart.formula_data[id] = data

          # Store range data locally to avoid lookup if seen again.
          seen_ranges[range] = data
        end
      end
    end

    #
    # Sort internal and user defined names in the same order as used by Excel.
    # This may not be strictly necessary but unsorted elements caused a lot of
    # issues in the the Spreadsheet::WriteExcel binary version. Also makes
    # comparison testing easier.
    #
    def sort_defined_names(names)
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
        else  # name_a == name_b
        # Secondary sort based on the sheet name.
          if sheet_a >= sheet_b
            1
          else
            -1
          end
        end
      end
    end

    # Used in the above sort routine to normalise the defined names. Removes any
    # leading '_xmln.' from internal names and lowercases the strings.
    def normalise_defined_name(name)
      name.sub(/^_xlnm./, '').downcase
    end

    # Used in the above sort routine to normalise the worksheet names for the
    # secondary sort. Removes leading quote and lowercases the strings.
    def normalise_sheet_name(name)
      name.sub(/^'/, '').downcase
    end

    #
    # Extract the named ranges from the sorted list of defined names. These are
    # used in the App.xml file.
    #
    def extract_named_ranges(defined_names)
      named_ranges = []

      defined_names.each do |defined_name|
        name, index, range = defined_name

        # Skip autoFilter ranges.
        next if name == '_xlnm._FilterDatabase'

        # We are only interested in defined names with ranges.
        if range =~ /^([^!]+)!/
          sheet_name = $1

          # Match Print_Area and Print_Titles xlnm types.
          if name =~ /^_xlnm\.(.*)$/
            xlnm_type = $1
            name = sheet_name + '!' + xlnm_type
          elsif index != -1
            name = sheet_name + '!' + name
          end

          named_ranges << name
        end
      end

      named_ranges
    end

    #
    # Iterate through the worksheets and set up any chart or image drawings.
    #
    def prepare_drawings
      chart_ref_id = 0
      image_ref_id = 0
      drawing_id   = 0
      @worksheets.each do |sheet|
        chart_count = sheet.charts.size
        image_count = sheet.images.size
        next if chart_count + image_count == 0

        drawing_id += 1

        (0 .. chart_count - 1).each do |index|
          chart_ref_id += 1
          sheet.prepare_chart(index, chart_ref_id, drawing_id)
        end

        (0 .. image_count - 1).each do |index|
          filename = sheet.images[index][2]

          image_id, type, width, height, name = get_image_properties(filename)

          image_ref_id += 1

          sheet.prepare_image(index, image_ref_id, drawing_id, width, height, name, type)
        end

        drawing = sheet.drawing
        @drawings << drawing
      end

      @drawing_count = drawing_id
    end

    #
    # Convert a sheet name to its index. Return undef otherwise.
    #
    def get_sheet_index(sheetname)
      sheet_count = @sheetnames.size
      sheet_index = nil

      sheetname.sub!(/^'/, '')
      sheetname.sub!(/'$/, '')

      ( 0 .. sheet_count - 1 ).each do |i|
        sheet_index = i if sheetname == @sheetnames[i]
      end

      sheet_index
    end
  end
end
