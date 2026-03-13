# -*- coding: utf-8 -*-
# frozen_string_literal: true

###############################################################################
#
# Workbook
#
# The Workbook class acts as a facade coordinating workbook state,
# package preparation and XML generation.
#
# Responsibilities are delegated to specialized modules:
#
#   Initialization      - workbook setup and default state
#   WorkbookWriter      - workbook XML generation
#   PackagePreparation  - package assembly and workbook preparation
#   FormatPreparation   - format, font, border and fill preparation
#   ChartData           - chart cache data extraction and defined name helpers
#
###############################################################################

require 'write_xlsx/workbook/initialization'
require 'write_xlsx/workbook/workbook_writer'
require 'write_xlsx/workbook/package_preparation'
require 'write_xlsx/workbook/format_preparation'
require 'write_xlsx/workbook/chart_data'
require 'write_xlsx/chart'
require 'write_xlsx/chartsheet'
require 'write_xlsx/format'
require 'write_xlsx/formats'
require 'write_xlsx/image_property'
require 'write_xlsx/shape'
require 'write_xlsx/sheets'
require 'write_xlsx/utility/common'
require 'write_xlsx/utility/cell_reference'
require 'write_xlsx/utility/xml_primitives'
require 'write_xlsx/worksheet'
require 'write_xlsx/zip_file_utils'
require 'write_xlsx/package/xml_writer_simple'
require 'write_xlsx/package/packager'
require 'tmpdir'
require 'tempfile'
require 'digest/md5'

module Writexlsx
  OFFICE_URL     = 'http://schemas.microsoft.com/office/'   # :nodoc:
  MAX_URL_LENGTH = 2_079

  class Workbook
    include Writexlsx::Utility::Common
    include Writexlsx::Utility::CellReference
    include Writexlsx::Utility::XmlPrimitives
    include Initialization
    include PackagePreparation
    include FormatPreparation
    include ChartData

    attr_writer :firstsheet                     # :nodoc:
    attr_reader :palette                        # :nodoc:
    attr_reader :worksheets                     # :nodoc:
    attr_accessor :drawings                     # :nodoc:
    attr_reader :named_ranges                   # :nodoc:
    attr_reader :doc_properties                 # :nodoc:
    attr_reader :custom_properties              # :nodoc:
    attr_reader :image_types, :images           # :nodoc:
    attr_reader :shared_strings                 # :nodoc:
    attr_reader :vba_project                    # :nodoc:
    attr_reader :excel2003_style                # :nodoc:
    attr_reader :max_url_length                 # :nodoc:
    attr_reader :strings_to_urls                # :nodoc:
    attr_reader :read_only                      # :nodoc:
    attr_reader :embedded_image_indexes         # :nodec:
    attr_reader :embedded_images                # :nodoc:
    attr_reader :embedded_descriptions          # :nodoc:
    attr_writer :has_embedded_descriptions      # :nodoc:
    attr_accessor :charts                       # :nodoc:

    ###############################################################################
    #
    # Lifecycle
    #
    ###############################################################################

    def initialize(file, *option_params)
      options, default_formats = process_workbook_options(*option_params)

      setup_core_state(file, options, default_formats)
      setup_workbook_state(options)
      setup_format_state(default_formats)
      setup_shared_strings
      setup_embedded_assets
      setup_calculation_state
      setup_default_formats
      set_color_palette
    end

    def close
      # In case close() is called twice.
      return if @fileclosed

      @fileclosed = true
      store_workbook
    end

    ###############################################################################
    #
    # Workbook object creation API
    #
    ###############################################################################

    #
    # At least one worksheet should be added to a new workbook. A worksheet is used to write data into cells:
    #
    def add_worksheet(name = '')
      name = check_sheetname(name)
      worksheet = Worksheet.new(self, @worksheets.size, name)
      @worksheets << worksheet
      worksheet
    end

    #
    # This method is use to create a new chart either as a standalone worksheet
    # (the default) or as an embeddable object that can be inserted into
    # a worksheet via the insert_chart method.
    #
    def add_chart(params = {})
      # Type must be specified so we can create the required chart instance.
      type     = params[:type]
      embedded = params[:embedded]
      name     = params[:name]
      raise "Must define chart type in add_chart()" unless type

      chart = Chart.factory(type, params[:subtype])
      chart.palette = @palette

      # If the chart isn't embedded let the workbook control it.
      if ptrue?(embedded)
        chart.name = name if name

        # Set index to 0 so that the activate() and set_first_sheet() methods
        # point back to the first worksheet if used for embedded charts.
        chart.index = 0
        chart.set_embedded_config_data
      else
        # Check the worksheet name for non-embedded charts.
        sheetname  = check_chart_sheetname(name)
        chartsheet = Chartsheet.new(self, @worksheets.size, sheetname)
        chartsheet.chart = chart
        @worksheets << chartsheet
      end
      @charts << chart
      ptrue?(embedded) ? chart : chartsheet
    end

    #
    # The add_format method can be used to create new Format objects
    # which are used to apply formatting to a cell. You can either define
    # the properties at creation time via a hash of property values
    # or later via method calls.
    #
    #     format1 = workbook.add_format(property_hash) # Set properties at creation
    #     format2 = workbook.add_format                # Set properties later
    #
    def add_format(property_hash = {})
      properties = {}
      properties.update(font: 'Arial', size: 10, theme: -1) if @excel2003_style
      properties.update(property_hash)

      format = Format.new(@formats, properties)

      @formats.formats.push(format)    # Store format reference

      format
    end

    #
    # The +add_shape+ method can be used to create new shapes that may be
    # inserted into a worksheet.
    #
    def add_shape(properties = {})
      shape = Shape.new(properties)
      shape.palette = @palette

      @shapes ||= []
      @shapes << shape  # Store shape reference.
      shape
    end

    ###############################################################################
    #
    # Workbook configuration API
    #
    ###############################################################################

    #
    # Set the date system: false = 1900 (the default), true = 1904
    #
    def set_1904(mode = true)
      raise "set_1904() must be called before add_worksheet()" unless sheets.empty?

      @date_1904 = ptrue?(mode)
    end

    #
    # return date system. false = 1900, true = 1904
    #
    def get_1904
      @date_1904
    end

    def set_tempdir(dir)
      @tempdir = dir.dup
    end

    #
    # Create a defined name in Excel. We handle global/workbook level names and
    # local/worksheet names.
    #
    def define_name(name, formula)
      sheet_index = nil
      sheetname   = ''

      # Local defined names are formatted like "Sheet1!name".
      if name =~ /^(.*)!(.*)$/
        sheetname   = ::Regexp.last_match(1)
        name        = ::Regexp.last_match(2)
        sheet_index = @worksheets.index_by_name(sheetname)
      else
        sheet_index = -1   # Use -1 to indicate global names.
      end

      # Raise if the sheet index wasn't found.
      raise "Unknown sheet name #{sheetname} in defined_name()" unless sheet_index

      # Raise if the name contains invalid chars as defined by Excel help.
      # Refer to the following to see Excel's syntax rules for defined names:
      # http://office.microsoft.com/en-001/excel-help/define-and-use-names-in-formulas-HA010147120.aspx#BMsyntax_rules_for_names
      #
      raise "Invalid characters in name '#{name}' used in defined_name()" if name =~ /\A[-0-9 !"#$%&'()*+,.:;<=>?@\[\]\^`{}~]/ || name =~ /.+[- !"#$%&'()*+,\\:;<=>?@\[\]\^`{}~]/

      # Raise if the name looks like a cell name.
      raise "Invalid name '#{name}' looks like a cell name in defined_name()" if name =~ /^[a-zA-Z][a-zA-Z]?[a-dA-D]?[0-9]+$/

      # Raise if the name looks like a R1C1
      raise "Invalid name '#{name}' like a RC cell ref in defined_name()" if name =~ /\A[rcRC]\Z/ || name =~ /\A[rcRC]\d+[rcRC]\d+\Z/

      @defined_names.push([name, sheet_index, formula.sub(/^=/, '')])
    end

    #
    # Set the workbook size.
    #
    def set_size(width = nil, height = nil)
      @window_width = if ptrue?(width)
                        # Convert to twips at 96 dpi.
                        width.to_i * 1440 / 96
                      else
                        16095
                      end

      @window_height = if ptrue?(height)
                         # Convert to twips at 96 dpi.
                         height.to_i * 1440 / 96
                       else
                         9660
                       end
    end

    #
    # Set the ratio of space for worksheet tabs.
    #
    def set_tab_ratio(tab_ratio = nil)
      return unless tab_ratio

      if tab_ratio < 0 || tab_ratio > 100
        raise "Tab ratio outside range: 0 <= zoom <= 100"
      else
        @tab_ratio = (tab_ratio * 10).to_i
      end
    end

    #
    # The set_properties method can be used to set the document properties
    # of the Excel file created by WriteXLSX. These properties are visible
    # when you use the Office Button -> Prepare -> Properties option in Excel
    # and are also available to external applications that read or index windows
    # files.
    #
    def set_properties(params)
      # Ignore if no args were passed.
      return -1 if params.empty?

      # List of valid input parameters.
      valid = {
        title:          1,
        subject:        1,
        author:         1,
        keywords:       1,
        comments:       1,
        last_author:    1,
        created:        1,
        category:       1,
        manager:        1,
        company:        1,
        status:         1,
        hyperlink_base: 1
      }

      # Check for valid input parameters.
      params.each_key do |key|
        return -1 unless valid.has_key?(key)
      end

      # Set the creation time unless specified by the user.
      params[:created] = @createtime unless params.has_key?(:created)

      @doc_properties = params.dup
    end

    #
    # Set a user defined custom document property.
    #
    def set_custom_property(name, value, type = nil)
      # Valid types.
      valid_type = {
        'text'       => 1,
        'date'       => 1,
        'number'     => 1,
        'number_int' => 1,
        'bool'       => 1
      }

      raise "The name and value parameters must be defined in set_custom_property()" if !name || (type != 'bool' && !value)

      # Determine the type for strings and numbers if it hasn't been specified.
      unless ptrue?(type)
        type = if value =~ /^\d+$/
                 'number_int'
               elsif value =~
                     /^([+-]?)(?=[0-9]|\.[0-9])[0-9]*(\.[0-9]*)?([Ee]([+-]?[0-9]+))?$/
                 'number'
               else
                 'text'
               end
      end

      # Check for valid validation types.
      raise "Unknown custom type '$type' in set_custom_property()" unless valid_type[type]

      #  Check for strings longer than Excel's limit of 255 chars.
      raise "Length of text custom value '$value' exceeds Excel's limit of 255 in set_custom_property()" if type == 'text' && value.length > 255

      if type == 'bool'
        value = value ? 1 : 0
      end

      @custom_properties << [name, value, type]
    end

    #
    # The add_vba_project method can be used to add macros or functions to an
    # WriteXLSX file using a binary VBA project file that has been extracted
    # from an existing Excel xlsm file.
    #
    def add_vba_project(vba_project)
      @vba_project = vba_project
    end

    #
    # Set the VBA name for the workbook.
    #
    def set_vba_name(vba_codename = nil)
      @vba_codename = vba_codename || 'ThisWorkbook'
    end

    #
    # Set the Excel "Read-only recommended" save option.
    #
    def read_only_recommended
      @read_only = 2
    end

    #
    # set_calc_mode()
    #
    # Set the Excel caclcuation mode for the workbook.
    #
    def set_calc_mode(mode, calc_id = nil)
      @calc_mode = mode || 'auto'

      if mode == 'manual'
        @calc_on_load = false
      elsif mode == 'auto_except_tables'
        @calc_mode = 'autoNoTable'
      end

      @calc_id = calc_id if calc_id
    end

    #
    # Change the RGB components of the elements in the colour palette.
    #
    def set_custom_color(index, red = 0, green = 0, blue = 0)
      # Match a HTML #xxyyzz style parameter
      if red.to_s =~ /^#(\w\w)(\w\w)(\w\w)/
        red   = ::Regexp.last_match(1).hex
        green = ::Regexp.last_match(2).hex
        blue  = ::Regexp.last_match(3).hex
      end

      # Check that the colour index is the right range
      raise "Color index #{index} outside range: 8 <= index <= 64" if index < 8 || index > 64

      # Check that the colour components are in the right range
      if (red   < 0 || red   > 255) ||
         (green < 0 || green > 255) ||
         (blue  < 0 || blue  > 255)
        raise "Color component outside range: 0 <= color <= 255"
      end

      index -= 8       # Adjust colour index (wingless dragonfly)

      # Set the RGB value
      @palette[index] = [red, green, blue]

      # Store the custome colors for the style.xml file.
      @custom_colors << sprintf("FF%02X%02X%02X", red, green, blue)

      index + 8
    end

    ###############################################################################
    #
    # Workbook accessors and lookup
    #
    ###############################################################################

    #
    # get array of Worksheet objects
    #
    # :call-seq:
    #   sheets              -> array of all Wordsheet object
    #   sheets(1, 3, 4)     -> array of spcified Worksheet object.
    #
    def sheets(*args)
      if args.empty?
        @worksheets
      else
        args.collect { |i| @worksheets[i] }
      end
    end

    #
    # Return a worksheet object in the workbook using the sheetname.
    #
    def worksheet_by_name(sheetname = nil)
      sheets.select { |s| s.name == sheetname }.first
    end
    alias get_worksheet_by_name worksheet_by_name

    #
    # user must not use. it is internal method.
    #
    def set_xml_writer(filename)  # :nodoc:
      @writer.set_xml_writer(filename)
    end

    #
    # user must not use. it is internal method.
    #
    def xml_str  # :nodoc:
      @writer.string
    end

    #
    # Get the default url format used when a user defined format isn't specified
    # with write_url(). The format is the hyperlink style defined by Excel for the
    # default theme.
    #
    attr_reader :default_url_format
    alias get_default_url_format default_url_format

    attr_writer :activesheet

    attr_reader :writer

    ###############################################################################
    #
    # Workbook state queries
    #
    ###############################################################################

    def date_1904? # :nodoc:
      @date_1904 ||= false
      !!@date_1904
    end

    def has_dynamic_functions?
      @has_dynamic_functions
    end

    #
    # Add a string to the shared string table, if it isn't already there, and
    # return the string index.
    #
    EMPTY_HASH = {}.freeze

    def shared_string_index(str) # :nodoc:
      @shared_strings.index(str, EMPTY_HASH)
    end

    def str_unique   # :nodoc:
      @shared_strings.unique_count
    end

    def shared_strings_empty?  # :nodoc:
      @shared_strings.empty?
    end

    def chartsheet_count
      @worksheets.chartsheet_count
    end

    def non_chartsheet_count
      @worksheets.worksheets.count
    end

    def style_properties
      [
        @xf_formats,
        @palette,
        @font_count,
        @num_formats,
        @border_count,
        @fill_count,
        @custom_colors,
        @dxf_formats,
        @has_comments
      ]
    end

    def num_vml_files
      @worksheets.select { |sheet| sheet.has_vml? || sheet.has_header_vml? }.count
    end

    def num_comment_files
      @worksheets.select { |sheet| sheet.has_comments? }.count
    end

    def chartsheets
      @worksheets.chartsheets
    end

    def non_chartsheets
      @worksheets.worksheets
    end

    def firstsheet # :nodoc:
      @firstsheet ||= 0
    end

    def activesheet # :nodoc:
      @activesheet ||= 0
    end

    def has_metadata?
      @has_metadata
    end

    def has_embedded_images?
      @has_embedded_images
    end

    def has_embedded_descriptions?
      @has_embedded_descriptions
    end

    #
    # Store the image types (PNG/JPEG/etc) used in the workbook to use in these
    # Content_Types file.
    #
    def store_image_types(type)
      case type
      when 'png'
        @image_types[:png] = 1
      when 'jpeg'
        @image_types[:jpeg] = 1
      when 'gif'
        @image_types[:gif] = 1
      when 'bmp'
        @image_types[:bmp] = 1
      end
    end

    ###############################################################################
    #
    # private helpers
    #
    ###############################################################################
    private

    ###############################################################################
    #
    # Worksheet and chart naming helpers
    #
    ###############################################################################

    #
    # Check for valid worksheet names. We check the length, if it contains any
    # invalid characters and if the name is unique in the workbook.
    #
    def check_sheetname(name) # :nodoc:
      @worksheets.make_and_check_sheet_chart_name(:sheet, name)
    end

    def check_chart_sheetname(name)
      @worksheets.make_and_check_sheet_chart_name(:chart, name)
    end

    ###############################################################################
    #
    # Internal utility helpers
    #
    ###############################################################################

    # for test
    def defined_names # :nodoc:
      @defined_names ||= []
    end

    def zip_entry_for_part(part)
      Zip::Entry.new("", part)
    end
  end
end
