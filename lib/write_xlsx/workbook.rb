# -*- coding: utf-8 -*-
require 'write_xlsx/package/xml_writer_simple'
require 'write_xlsx/package/packager'
require 'write_xlsx/sheets'
require 'write_xlsx/worksheet'
require 'write_xlsx/chartsheet'
require 'write_xlsx/formats'
require 'write_xlsx/format'
require 'write_xlsx/shape'
require 'write_xlsx/utility'
require 'write_xlsx/chart'
require 'write_xlsx/zip_file_utils'
require 'tmpdir'
require 'tempfile'
require 'digest/md5'

module Writexlsx

  OFFICE_URL = 'http://schemas.microsoft.com/office/'   # :nodoc:

  # The WriteXLSX provides an object oriented interface to a new Excel workbook.
  # The following methods are available through a new workbook.
  #
  # * new[#method-c-new]
  # * add_worksheet[#method-i-add_worksheet]
  # * add_format[#method-i-add_format]
  # * add_chart[#method-i-add_chart]
  # * add_shape[#method-i-add_shape]
  # * add_vba_project[#method-i-add_vba_project]
  # * close[#method-i-close]
  # * set_properties[#method-i-set_properties]
  # * define_name[#method-i-define_name]
  # * set_custom_color[#method-i-set_custom_color]
  # * sheets[#method-i-sheets]
  # * set_1904[#method-i-set_1904]
  #
  class Workbook

    include Writexlsx::Utility

    attr_writer :firstsheet  # :nodoc:
    attr_reader :palette  # :nodoc:
    attr_reader :worksheets, :charts, :drawings  # :nodoc:
    attr_reader :named_ranges   # :nodoc:
    attr_reader :doc_properties  # :nodoc:
    attr_reader :image_types, :images  # :nodoc:
    attr_reader :shared_strings  # :nodoc:
    attr_reader :vba_project  # :nodoc:
    attr_reader :excel2003_style # :nodoc:
    attr_reader :strings_to_urls # :nodoc:
    #
    # A new Excel workbook is created using the +new+ constructor
    # which accepts either a filename or an IO object as a parameter.
    # The following example creates a new Excel file based on a filename:
    #
    #     workbook  = WriteXLSX.new('filename.xlsx')
    #     worksheet = workbook.add_worksheet
    #     worksheet.write(0, 0, 'Hi Excel!')
    #     workbook.close
    #
    # Here are some other examples of using +new+ with filenames:
    #
    #     workbook1 = WriteXLSX.new(filename)
    #     workbook2 = WriteXLSX.new('/tmp/filename.xlsx')
    #     workbook3 = WriteXLSX.new("c:\\tmp\\filename.xlsx")
    #     workbook4 = WriteXLSX.new('c:\tmp\filename.xlsx')
    #
    # The last two examples demonstrates how to create a file on DOS or Windows
    # where it is necessary to either escape the directory separator \
    # or to use single quotes to ensure that it isn't interpolated.
    #
    # It is recommended that the filename uses the extension .xlsx
    # rather than .xls since the latter causes an Excel warning
    # when used with the XLSX format.
    #
    # The +new+ constructor returns a WriteXLSX object that you can use to
    # add worksheets and store data.
    #
    # You can also pass a valid IO object to the +new+ constructor.
    #
    #     xlsx = StringIO.new
    #     workbook = WriteXLSX.new(xlsx)
    #     ....
    #     workbook.close
    #     # you can get XLSX binary data as xlsx.string
    #
    # And you can pass default_formats parameter like this:
    #
    #     formats = { :font => 'Arial', :size => 10.5 }
    #     workbook = WriteXLSX.new('file.xlsx', formats)
    #
    def initialize(file, *option_params)
      options, default_formats = process_workbook_options(*option_params)
      @writer = Package::XMLWriterSimple.new

      @file                = file
      @tempdir = options[:tempdir] ||
        File.join(Dir.tmpdir, Digest::MD5.hexdigest("#{Time.now.to_f.to_s}-#{Process.pid}"))
      @date_1904           = options[:date_1904] || false
      @activesheet         = 0
      @firstsheet          = 0
      @selected            = 0
      @fileclosed          = false
      @worksheets          = Sheets.new
      @charts              = []
      @drawings            = []
      @formats             = Formats.new
      @xf_formats          = []
      @dxf_formats         = []
      @font_count          = 0
      @num_format_count    = 0
      @defined_names       = []
      @named_ranges        = []
      @custom_colors       = []
      @doc_properties      = {}
      @local_time          = Time.now
      @optimization        = options[:optimization] || 0
      @x_window            = 240
      @y_window            = 15
      @window_width        = 16095
      @window_height       = 9660
      @tab_ratio           = 500
      @excel2003_style     = options[:excel2003_style] || false
      @table_count         = 0
      @image_types         = {}
      @images              = []
      @strings_to_urls     = (options[:strings_to_urls].nil? || options[:strings_to_urls]) ? true : false

      # Structures for the shared strings data.
      @shared_strings = Package::SharedStrings.new

      # Formula calculation default settings.
      @calc_id             = 124519
      @calc_mode           = 'auto'
      @calc_on_load        = true

      if @excel2003_style
        add_format(default_formats
                     .merge(:xf_index => 0, :font_family => 0, :font => 'Arial', :size => 10, :theme => -1))
      else
        add_format(default_formats.merge(:xf_index => 0))
      end
      set_color_palette
    end

    #
    # The close method is used to close an Excel file.
    #
    # An explicit close is required if the file must be closed prior to performing
    # some external action on it such as copying it, reading its size or attaching
    # it to an email.
    #
    # In general, if you create a file with a size of 0 bytes or you fail to create
    # a file you need to call close.
    #
    def close
      # In case close() is called twice.
      return if @fileclosed

      @fileclosed = true
      store_workbook
    end

    #
    # get array of Worksheet objects
    #
    # :call-seq:
    #   sheets              -> array of all Wordsheet object
    #   sheets(1, 3, 4)     -> array of spcified Worksheet object.
    #
    # The sheets method returns a array, or a sliced array, of the worksheets
    # in a workbook.
    #
    # If no arguments are passed the method returns a list of all the worksheets
    # in the workbook. This is useful if you want to repeat an operation on each
    # worksheet:
    #
    #     workbook.sheets.each do |worksheet|
    #        print worksheet.get_name
    #     end
    #
    # You can also specify a slice list to return one or more worksheet objects:
    #
    #     worksheet = workbook.sheets(0)
    #     worksheet.write('A1', 'Hello')
    #
    # you can write the above example as:
    #
    #     workbook.sheets(0).write('A1', 'Hello')
    #
    # The following example returns the first and last worksheet in a workbook:
    #
    #     workbook.sheets(0, -1).each do |sheet|
    #        # Do something
    #     end
    #
    def sheets(*args)
      if args.empty?
        @worksheets
      else
        args.collect{|i| @worksheets[i] }
      end
    end

    #
    # Set the date system: false = 1900 (the default), true = 1904
    #
    # Excel stores dates as real numbers where the integer part stores
    # the number of days since the epoch and the fractional part stores
    # the percentage of the day. The epoch can be either 1900 or 1904.
    # Excel for Windows uses 1900 and Excel for Macintosh uses 1904.
    # However, Excel on either platform will convert automatically between
    # one system and the other.
    #
    # WriteXLSX stores dates in the 1900 format by default. If you wish to
    # change this you can call the set_1904 workbook method.
    # You can query the current value by calling the get_1904 workbook method.
    # This returns false for 1900 and true for 1904.
    #
    # In general you probably won't need to use set_1904.
    #
    def set_1904(mode = true)
      unless sheets.empty?
        raise "set_1904() must be called before add_worksheet()"
      end
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
    # user must not use. it is internal method.
    #
    def set_xml_writer(filename)  #:nodoc:
      @writer.set_xml_writer(filename)
    end

    #
    # user must not use. it is internal method.
    #
    def xml_str  #:nodoc:
      @writer.string
    end

    #
    # user must not use. it is internal method.
    #
    def assemble_xml_file  #:nodoc:
      return unless @writer

      # Prepare format object for passing to Style.rb.
      prepare_format_properties

      write_xml_declaration do

        # Write the root workbook element.
        write_workbook do

          # Write the XLSX file version.
          write_file_version

          # Write the workbook properties.
          write_workbook_pr

          # Write the workbook view properties.
          write_book_views

          # Write the worksheet names and ids.
          @worksheets.write_sheets(@writer)

          # Write the workbook defined names.
          write_defined_names

          # Write the workbook calculation properties.
          write_calc_pr

          # Write the workbook extension storage.
          #write_ext_lst
        end
      end
    end

    #
    # At least one worksheet should be added to a new workbook. A worksheet is used to write data into cells:
    #
    #     worksheet1 = workbook.add_worksheet               # Sheet1
    #     worksheet2 = workbook.add_worksheet('Foglio2')    # Foglio2
    #     worksheet3 = workbook.add_worksheet('Data')       # Data
    #     worksheet4 = workbook.add_worksheet               # Sheet4
    # If name is not specified the default Excel convention will be followed, i.e. Sheet1, Sheet2, etc.
    #
    # The worksheet name must be a valid Excel worksheet name,
    # i.e. it cannot contain any of the following characters,
    #     [ ] : * ? / \
    #
    # and it must be less than 32 characters.
    # In addition, you cannot use the same, case insensitive,
    # sheetname for more than one worksheet.
    #
    def add_worksheet(name = '')
      name  = check_sheetname(name)
      worksheet = Worksheet.new(self, @worksheets.size, name)
      @worksheets << worksheet
      worksheet
    end

    #
    # This method is use to create a new chart either as a standalone worksheet
    # (the default) or as an embeddable object that can be inserted into
    # a worksheet via the
    # {Worksheet#insert_chart}[Worksheet.html#method-i-insert_chart] method.
    #
    #     chart = workbook.add_chart(:type => 'column')
    #
    # The properties that can be set are:
    #
    #     :type     (required)
    #     :subtype  (optional)
    #     :name     (optional)
    #     :embedded (optional)
    #
    # === :type
    #
    # This is a required parameter.
    # It defines the type of chart that will be created.
    #
    #     chart = workbook.add_chart(:type => 'line')
    #
    # The available types are:
    #
    #     area
    #     bar
    #     column
    #     line
    #     pie
    #     scatter
    #     stock
    #
    # === :subtype
    #
    # Used to define a chart subtype where available.
    #
    #     chart = workbook.add_chart(:type => 'bar', :subtype => 'stacked')
    #
    # Currently only Bar and Column charts support subtypes
    # (stacked and percent_stacked). See the documentation for those chart
    # types.
    #
    # === :name
    #
    # Set the name for the chart sheet. The name property is optional and
    # if it isn't supplied will default to Chart1 .. n. The name must be
    # a valid Excel worksheet name. See add_worksheet
    # for more details on valid sheet names. The name property can be
    # omitted for embedded charts.
    #
    #     chart = workbook.add_chart(:type => 'line', :name => 'Results Chart')
    #
    # === :embedded
    #
    # Specifies that the Chart object will be inserted in a worksheet
    # via the {Worksheet#insert_chart}[Worksheet.html#insert_chart] method.
    # It is an error to try insert a Chart that doesn't have this flag set.
    #
    #     chart = workbook.add_chart(:type => 'line', :embedded => 1)
    #
    #     # Configure the chart.
    #     ...
    #
    #     # Insert the chart into the a worksheet.
    #     worksheet.insert_chart('E2', chart)
    #
    # See Chart[Chart.html] for details on how to configure the chart object
    # once it is created. See also the chart_*.rb programs in the examples
    # directory of the distro.
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
        chartsheet.chart   = chart
        @worksheets << chartsheet
      end
      @charts << chart
      ptrue?(embedded) ? chart : chartsheet
    end

    #
    # The +add_format+ method can be used to create new Format objects
    # which are used to apply formatting to a cell. You can either define
    # the properties at creation time via a hash of property values
    # or later via method calls.
    #
    #     format1 = workbook.add_format(property_hash) # Set properties at creation
    #     format2 = workbook.add_format                # Set properties later
    #
    # See the {Format Class's rdoc}[Format.html] for more details about
    # Format properties and how to set them.
    #
    def add_format(property_hash = {})
      properties = {}
      if @excel2003_style
        properties.update(:font => 'Arial', :size => 10, :theme => -1)
      end
      properties.update(property_hash)

      format = Format.new(@formats, properties)

      @formats.formats.push(format)    # Store format reference

      format
    end

    #
    # The +add_shape+ method can be used to create new shapes that may be
    # inserted into a worksheet.
    #
    # You can either define the properties at creation time via a hash of
    # property values or later via method calls.
    #
    #     # Set properties at creation.
    #     plus  = workbook.add_shape(
    #               :type   => 'plus',
    #               :id     => 3,
    #               :width  => pw,
    #               :height => ph
    #             )
    #
    #     # Default rectangle shape. Set properties later.
    #     rect  = workbook.add_shape
    #
    # See also the shape*.rb programs in the examples directory of the distro.
    #
    # === Shape Properties
    #
    # Any shape property can be queried or modified by [ ] like hash.
    #
    #   ellipse = workbook.add_shape(properties)
    #   ellipse[:type] = 'cross'    # No longer an ellipse !
    #   type = ellipse[:type]       # Find out what it really is.
    #
    # The properties of a shape object that can be defined via add_shape are
    # shown below.
    #
    # ===:name
    #
    # Defines the name of the shape. This is an optional property and the shape
    # will be given a default name if not supplied. The name is generally only
    # used by Excel Macros to refer to the object.
    #
    # ===:type
    #
    # Defines the type of the object such as +:rect+, +:ellipse+ OR +:triangle+.
    #
    #   ellipse = workbook.add_shape(:type => :ellipse)
    #
    # The default type is +:rect+.
    #
    # The full list of available shapes is shown below.
    #
    # See also the shape_all.rb program in the examples directory of the distro.
    # It creates an example workbook with all supported shapes labelled with
    # their shape names.
    #
    # === Basic Shapes
    #
    #    blockArc              can            chevron       cube          decagon
    #    diamond               dodecagon      donut         ellipse       funnel
    #    gear6                 gear9          heart         heptagon      hexagon
    #    homePlate             lightningBolt  line          lineInv       moon
    #    nonIsoscelesTrapezoid noSmoking      octagon       parallelogram pentagon
    #    pie                   pieWedge       plaque        rect          round1Rect
    #    round2DiagRect        round2SameRect roundRect     rtTriangle    smileyFace
    #    snip1Rect             snip2DiagRect  snip2SameRect snipRoundRect star10
    #    star12                star16         star24        star32        star4
    #    star5                 star6          star7         star8         sun
    #    teardrop              trapezoid      triangle
    #
    # === Arrow Shapes
    #
    #    bentArrow        bentUpArrow       circularArrow     curvedDownArrow
    #    curvedLeftArrow  curvedRightArrow  curvedUpArrow     downArrow
    #    leftArrow        leftCircularArrow leftRightArrow    leftRightCircularArrow
    #    leftRightUpArrow leftUpArrow       notchedRightArrow quadArrow
    #    rightArrow       stripedRightArrow swooshArrow       upArrow
    #    upDownArrow      uturnArrow
    #
    # === Connector Shapes
    #
    #    bentConnector2   bentConnector3   bentConnector4
    #    bentConnector5   curvedConnector2 curvedConnector3
    #    curvedConnector4 curvedConnector5 straightConnector1
    #
    # === Callout Shapes
    #
    #    accentBorderCallout1  accentBorderCallout2  accentBorderCallout3
    #    accentCallout1        accentCallout2        accentCallout3
    #    borderCallout1        borderCallout2        borderCallout3
    #    callout1              callout2              callout3
    #    cloudCallout          downArrowCallout      leftArrowCallout
    #    leftRightArrowCallout quadArrowCallout      rightArrowCallout
    #    upArrowCallout        upDownArrowCallout    wedgeEllipseCallout
    #    wedgeRectCallout      wedgeRoundRectCallout
    #
    # === Flow Chart Shapes
    #
    #    flowChartAlternateProcess  flowChartCollate        flowChartConnector
    #    flowChartDecision          flowChartDelay          flowChartDisplay
    #    flowChartDocument          flowChartExtract        flowChartInputOutput
    #    flowChartInternalStorage   flowChartMagneticDisk   flowChartMagneticDrum
    #    flowChartMagneticTape      flowChartManualInput    flowChartManualOperation
    #    flowChartMerge             flowChartMultidocument  flowChartOfflineStorage
    #    flowChartOffpageConnector  flowChartOnlineStorage  flowChartOr
    #    flowChartPredefinedProcess flowChartPreparation    flowChartProcess
    #    flowChartPunchedCard       flowChartPunchedTape    flowChartSort
    #    flowChartSummingJunction   flowChartTerminator
    #
    # === Action Shapes
    #
    #    actionButtonBackPrevious actionButtonBeginning actionButtonBlank
    #    actionButtonDocument     actionButtonEnd       actionButtonForwardNext
    #    actionButtonHelp         actionButtonHome      actionButtonInformation
    #    actionButtonMovie        actionButtonReturn    actionButtonSound
    #
    # === Chart Shapes
    #
    # Not to be confused with Excel Charts.
    #
    #    chartPlus chartStar chartX
    #
    # === Math Shapes
    #
    #    mathDivide mathEqual mathMinus mathMultiply mathNotEqual mathPlus
    #
    # === Starts and Banners
    #
    #    arc            bevel          bracePair  bracketPair chord
    #    cloud          corner         diagStripe doubleWave  ellipseRibbon
    #    ellipseRibbon2 foldedCorner   frame      halfFrame   horizontalScroll
    #    irregularSeal1 irregularSeal2 leftBrace  leftBracket leftRightRibbon
    #    plus           ribbon         ribbon2    rightBrace  rightBracket
    #    verticalScroll wave
    #
    # === Tab Shapes
    #
    #    cornerTabs plaqueTabs squareTabs
    #
    # === :text
    #
    # This property is used to make the shape act like a text box.
    #
    #    rect = workbook.add_shape(:type => 'rect', :text => "Hello \nWorld")
    #
    # The Text is super-imposed over the shape. The text can be wrapped using
    # the newline character \n.
    #
    # === :id
    #
    # Identification number for internal identification. This number will be
    # auto-assigned, if not assigned, or if it is a duplicate.
    #
    # === :format
    #
    # Workbook format for decorating the shape horizontally and/or vertically.
    #
    # === :rotation
    #
    # Shape rotation, in degrees, from 0 to 360
    #
    # === :line, :fill
    #
    # Shape color for the outline and fill.
    # Colors may be specified  as a color index, or in RGB format, i.e. AA00FF.
    #
    # See COULOURS IN EXCEL in the main documentation for more information.
    #
    # === :link_type
    #
    # Line type for shape outline. The default is solid.
    # The list of possible values is:
    #
    #    dash, sysDot, dashDot, lgDash, lgDashDot, lgDashDotDot, solid
    #
    # === :valign, :align
    #
    # Text alignment within the shape.
    #
    # Vertical alignment can be:
    #
    #    Setting     Meaning
    #    =======     =======
    #    t           Top
    #    ctr         Centre
    #    b           Bottom
    #
    # Horizontal alignment can be:
    #
    #    Setting     Meaning
    #    =======     =======
    #    l           Left
    #    r           Right
    #    ctr         Centre
    #    just        Justified
    #
    # The default is to center both horizontally and vertically.
    #
    # === :scale_x, :scale_y
    #
    # Scale factor in x and y dimension, for scaling the shape width and
    # height. The default value is 1.
    #
    # Scaling may be set on the shape object or via insert_shape.
    #
    # === :adjustments
    #
    # Adjustment of shape vertices. Most shapes do not use this. For some
    # shapes, there is a single adjustment to modify the geometry.
    # For instance, the plus shape has one adjustment to control the width
    # of the spokes.
    #
    # Connectors can have a number of adjustments to control the shape
    # routing. Typically, a connector will have 3 to 5 handles for routing
    # the shape. The adjustment is in percent of the distance from the
    # starting shape to the ending shape, alternating between the x and y
    # dimension. Adjustments may be negative, to route the shape away
    # from the endpoint.
    #
    # === :stencil
    #
    # Shapes work in stencil mode by default. That is, once a shape is
    # inserted, its connection is separated from its master.
    # The master shape may be modified after an instance is inserted,
    # and only subsequent insertions will show the modifications.
    #
    # This is helpful for Org charts, where an employee shape may be
    # created once, and then the text of the shape is modified for each
    # employee.
    #
    # The insert_shape method returns a reference to the inserted
    # shape (the child).
    #
    # Stencil mode can be turned off, allowing for shape(s) to be
    # modified after insertion. In this case the insert_shape() method
    # returns a reference to the inserted shape (the master).
    # This is not very useful for inserting multiple shapes,
    # since the x/y coordinates also gets modified.
    #
    def add_shape(properties = {})
      shape = Shape.new(properties)
      shape.palette = @palette

      @shapes ||= []
      @shapes << shape  #Store shape reference.
      shape
    end

    #
    # Create a defined name in Excel. We handle global/workbook level names and
    # local/worksheet names.
    #
    # This method is used to defined a name that can be used to represent
    # a value, a single cell or a range of cells in a workbook.
    #
    # For example to set a global/workbook name:
    #
    #   # Global/workbook names.
    #   workbook.define_name('Exchange_rate', '=0.96')
    #   workbook.define_name('Sales',         '=Sheet1!$G$1:$H$10')
    #
    # It is also possible to define a local/worksheet name by prefixing the name
    # with the sheet name using the syntax +sheetname!definedname+:
    #
    #   # Local/worksheet name.
    #   workbook.define_name('Sheet2!Sales',  '=Sheet2!$G$1:$G$10')
    #
    # If the sheet name contains spaces or special characters
    # you must enclose it in single quotes like in Excel:
    #
    #   workbook.define_name("'New Data'!Sales",  '=Sheet2!$G$1:$G$10')
    #
    # See the defined_name.rb program in the examples dir of the distro.
    #
    # Refer to the following to see Excel's syntax rules for defined names:
    # <http://office.microsoft.com/en-001/excel-help/define-and-use-names-in-formulas-HA010147120.aspx#BMsyntax_rules_for_names>
    #
    def define_name(name, formula)
      sheet_index = nil
      sheetname   = ''

      # Local defined names are formatted like "Sheet1!name".
      if name =~ /^(.*)!(.*)$/
        sheetname   = $1
        name        = $2
        sheet_index = @worksheets.index_by_name(sheetname)
      else
        sheet_index = -1   # Use -1 to indicate global names.
      end

      # Raise if the sheet index wasn't found.
      if !sheet_index
       raise "Unknown sheet name #{sheetname} in defined_name()\n"
      end

      # Raise if the name contains invalid chars as defined by Excel help.
      # Refer to the following to see Excel's syntax rules for defined names:
      # http://office.microsoft.com/en-001/excel-help/define-and-use-names-in-formulas-HA010147120.aspx#BMsyntax_rules_for_names
      #
      if name =~ /\A[-0-9 !"#\$%&'\(\)\*\+,\.:;<=>\?@\[\]\^`\{\}~]/ || name =~ /.+[- !"#\$%&'\(\)\*\+,\\:;<=>\?@\[\]\^`\{\}~]/
        raise "Invalid characters in name '#{name}' used in defined_name()\n"
      end

      # Raise if the name looks like a cell name.
      if name =~ %r(^[a-zA-Z][a-zA-Z]?[a-dA-D]?[0-9]+$)
        raise "Invalid name '#{name}' looks like a cell name in defined_name()\n"
      end

      # Raise if the name looks like a R1C1
      if name =~ /\A[rcRC]\Z/ || name =~ /\A[rcRC]\d+[rcRC]\d+\Z/
        raise "Invalid name '#{name}' like a RC cell ref in defined_name()\n"
      end

      @defined_names.push([ name, sheet_index, formula.sub(/^=/, '') ])
    end

    #
    # The set_properties method can be used to set the document properties
    # of the Excel file created by WriteXLSX. These properties are visible
    # when you use the Office Button -> Prepare -> Properties option in Excel
    # and are also available to external applications that read or index windows
    # files.
    #
    # The properties should be passed in hash format as follows:
    #
    #     workbook.set_properties(
    #       :title    => 'This is an example spreadsheet',
    #       :author   => 'Hideo NAKAMURA',
    #       :comments => 'Created with Ruby and WriteXLSX'
    #     )
    #
    # The properties that can be set are:
    #
    #     :title
    #     :subject
    #     :author
    #     :manager
    #     :company
    #     :category
    #     :keywords
    #     :comments
    #     :status
    #
    # See also the properties.rb program in the examples directory
    # of the distro.
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

    #
    # The add_vba_project method can be used to add macros or functions to an
    # WriteXLSX file using a binary VBA project file that has been extracted
    # from an existing Excel xlsm file.
    #
    #     workbook  = WriteXLSX.new('file.xlsm')
    #
    #     workbook.add_vba_project('./vbaProject.bin')
    #
    # The supplied +extract_vba+ utility can be used to extract the required
    # +vbaProject.bin+ file from an existing Excel file:
    #
    #     $ extract_vba file.xlsm
    #     Extracted 'vbaProject.bin' successfully
    #
    # Macros can be tied to buttons using the worksheet
    # {insert_button}[Worksheet.html#method-i-insert_button] method
    # (see the "WORKSHEET METHODS" section for details):
    #
    #     worksheet.insert_button('C2', { :macro => 'my_macro' })
    #
    # Note, Excel uses the file extension xlsm instead of xlsx for files that
    # contain macros. It is advisable to follow the same convention.
    #
    # See also the macros.rb example file.
    #
    def add_vba_project(vba_project)
      @vba_project = vba_project
    end

    #
    # Set the VBA name for the workbook.
    #
    def set_vba_name(vba_codename = nil)
      if vba_codename
        @vba_codename = vba_codename
      else
        @vba_codename = 'ThisWorkbook'
      end
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
    # The set_custom_color method can be used to override one of the built-in
    # palette values with a more suitable colour.
    #
    # The value for +index+ should be in the range 8..63,
    # see "COLOURS IN EXCEL".
    #
    # The default named colours use the following indices:
    #
    #      8   =>   black
    #      9   =>   white
    #     10   =>   red
    #     11   =>   lime
    #     12   =>   blue
    #     13   =>   yellow
    #     14   =>   magenta
    #     15   =>   cyan
    #     16   =>   brown
    #     17   =>   green
    #     18   =>   navy
    #     20   =>   purple
    #     22   =>   silver
    #     23   =>   gray
    #     33   =>   pink
    #     53   =>   orange
    #
    # A new colour is set using its RGB (red green blue) components. The +red+,
    # +green+ and +blue+ values must be in the range 0..255. You can determine
    # the required values in Excel using the Tools->Options->Colors->Modify
    # dialog.
    #
    # The set_custom_color workbook method can also be used with a HTML style
    # +#rrggbb+ hex value:
    #
    #     workbook.set_custom_color(40, 255,  102,  0   ) # Orange
    #     workbook.set_custom_color(40, 0xFF, 0x66, 0x00) # Same thing
    #     workbook.set_custom_color(40, '#FF6600'       ) # Same thing
    #
    #     font = workbook.add_format(:color => 40)   # Use the modified colour
    #
    # The return value from set_custom_color() is the index of the colour that
    # was changed:
    #
    #     ferrari = workbook.set_custom_color(40, 216, 12, 12)
    #
    #     format  = workbook.add_format(
    #                                 :bg_color => ferrari,
    #                                 :pattern  => 1,
    #                                 :border   => 1
    #                            )
    #
    # Note, In the XLSX format the color palette isn't actually confined to 53
    # unique colors. The WriteXLSX gem will be extended at a later stage to
    # support the newer, semi-infinite, palette.
    #
    def set_custom_color(index, red = 0, green = 0, blue = 0)
      # Match a HTML #xxyyzz style parameter
      if red =~ /^#(\w\w)(\w\w)(\w\w)/
        red   = $1.hex
        green = $2.hex
        blue  = $3.hex
      end

      # Check that the colour index is the right range
      if index < 8 || index > 64
        raise "Color index #{index} outside range: 8 <= index <= 64"
      end

      # Check that the colour components are in the right range
      if (red   < 0 || red   > 255) ||
         (green < 0 || green > 255) ||
         (blue  < 0 || blue  > 255)
        raise "Color component outside range: 0 <= color <= 255"
      end

      index -=8       # Adjust colour index (wingless dragonfly)

      # Set the RGB value
      @palette[index] = [red, green, blue]

      # Store the custome colors for the style.xml file.
      @custom_colors << sprintf("FF%02X%02X%02X", red, green, blue)

      index + 8
    end

    def activesheet=(worksheet) #:nodoc:
      @activesheet = worksheet
    end

    def writer #:nodoc:
      @writer
    end

    def date_1904? #:nodoc:
      @date_1904 ||= false
      !!@date_1904
    end

    #
    # Add a string to the shared string table, if it isn't already there, and
    # return the string index.
    #
    def shared_string_index(str, params = {}) #:nodoc:
      @shared_strings.index(str, params)
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
       @num_format_count,
       @border_count,
       @fill_count,
       @custom_colors,
       @dxf_formats
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

    def firstsheet #:nodoc:
      @firstsheet ||= 0
    end

    def activesheet #:nodoc:
      @activesheet ||= 0
    end

    private

    def filename
      setup_filename unless @filename
      @filename
    end

    def fileobj
      setup_filename unless @fileobj
      @fileobj
    end

    def setup_filename #:nodoc:
      if @file.respond_to?(:to_str) && @file != ''
        @filename = @file
        @fileobj  = nil
      elsif @file.respond_to?(:write)
        @filename = File.join(tempdir, Digest::MD5.hexdigest(Time.now.to_s) + '.xlsx.tmp')
        @fileobj  = @file
      else
        raise "'#{@file}' must be valid filename String of IO object."
      end
    end

    def tempdir
      @tempdir
    end

    #
    # Sets the colour palette to the Excel defaults.
    #
    def set_color_palette #:nodoc:
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
    # Check for valid worksheet names. We check the length, if it contains any
    # invalid characters and if the name is unique in the workbook.
    #
    def check_sheetname(name) #:nodoc:
      @worksheets.make_and_check_sheet_chart_name(:sheet, name)
    end

    def check_chart_sheetname(name)
      @worksheets.make_and_check_sheet_chart_name(:chart, name)
    end

    #
    # Convert a range formula such as Sheet1!$B$1:$B$5 into a sheet name and cell
    # range such as ( 'Sheet1', 0, 1, 4, 1 ).
    #
    def get_chart_range(range) #:nodoc:
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

    def write_workbook #:nodoc:
      schema  = 'http://schemas.openxmlformats.org'
      attributes = [
        ['xmlns',
         schema + '/spreadsheetml/2006/main'],
        ['xmlns:r',
         schema + '/officeDocument/2006/relationships']
      ]
      @writer.tag_elements('workbook', attributes) do
        yield
      end
    end

    def write_file_version #:nodoc:
      attributes = [
                    ['appName', 'xl'],
                    ['lastEdited', 4],
                    ['lowestEdited', 4],
                    ['rupBuild', 4505]
                   ]

      if @vba_project
        attributes << [:codeName, '{37E998C4-C9E5-D4B9-71C8-EB1FF731991C}']
      end

      @writer.empty_tag('fileVersion', attributes)
    end

    def write_workbook_pr #:nodoc:
      attributes = []
      attributes << ['codeName', @vba_codename]  if ptrue?(@vba_codename)
      attributes << ['date1904', 1]              if date_1904?
      attributes << ['defaultThemeVersion', 124226]
      @writer.empty_tag('workbookPr', attributes)
    end

    def write_book_views #:nodoc:
      @writer.tag_elements('bookViews') { write_workbook_view }
    end

    def write_workbook_view #:nodoc:
      attributes = [
                    ['xWindow',       @x_window],
                    ['yWindow',       @y_window],
                    ['windowWidth',   @window_width],
                    ['windowHeight',  @window_height]
                   ]
      if @tab_ratio != 500
        attributes << ['tabRatio', @tab_ratio]
      end
      if @firstsheet > 0
        attributes << ['firstSheet', @firstsheet + 1]
      end
      if @activesheet > 0
        attributes << ['activeTab', @activesheet]
      end
      @writer.empty_tag('workbookView', attributes)
    end

    def write_calc_pr #:nodoc:
      attributes = [ ['calcId', @calc_id] ]

      case @calc_mode
      when 'manual'
        attributes << ['calcMode', 'manual']
        attributes << ['calcOnSave', 0]
      when 'autoNoTable'
        attributes << ['calcMode', 'autoNoTable']
      end

      attributes << ['fullCalcOnLoad', 1] if @calc_on_load

      @writer.empty_tag('calcPr', attributes)
    end

    def write_ext_lst #:nodoc:
      @writer.tag_elements('extLst') { write_ext }
    end

    def write_ext #:nodoc:
      attributes = [
        ['xmlns:mx', "#{OFFICE_URL}mac/excel/2008/main"],
        ['uri', uri]
      ]
      @writer.tag_elements('ext', attributes) { write_mx_arch_id }
    end

    def write_mx_arch_id #:nodoc:
      @writer.empty_tag('mx:ArchID', ['Flags', 2])
    end

    def write_defined_names #:nodoc:
      return unless ptrue?(@defined_names)
      @writer.tag_elements('definedNames') do
        @defined_names.each { |defined_name| write_defined_name(defined_name) }
      end
    end

    def write_defined_name(defined_name) #:nodoc:
      name, id, range, hidden = defined_name

      attributes = [ ['name', name] ]
      attributes << ['localSheetId', "#{id}"] unless id == -1
      attributes << ['hidden',       '1']     if hidden

      @writer.data_element('definedName', range, attributes)
    end

    def write_io(str) #:nodoc:
      @writer << str
      str
    end

    # for test
    def defined_names #:nodoc:
      @defined_names ||= []
    end

    #
    # Assemble worksheets into a workbook.
    #
    def store_workbook #:nodoc:
      # Add a default worksheet if non have been added.
      add_worksheet if @worksheets.empty?

      # Ensure that at least one worksheet has been selected.
      @worksheets.visible_first.select if @activesheet == 0

      # Set the active sheet.
      @activesheet = @worksheets.visible_first.index if @activesheet == 0
      @worksheets[@activesheet].activate

      # Prepare the worksheet VML elements such as comments and buttons.
      prepare_vml_objects
      # Set the defined names for the worksheets such as Print Titles.
      prepare_defined_names
      # Prepare the drawings, charts and images.
      prepare_drawings
      # Add cached data to charts.
      add_chart_data

      # Prepare the worksheet tables.
      prepare_tables

      # Package the workbook.
      packager = Package::Packager.new(self)
      packager.set_package_dir(tempdir)
      packager.create_package

      # Free up the Packager object.
      packager = nil

      # Store the xlsx component files with the temp dir name removed.
      ZipFileUtils.zip("#{tempdir}", filename)

      IO.copy_stream(filename, fileobj) if fileobj
      Writexlsx::Utility.delete_files(tempdir)
    end

    def write_parts(zip)
      parts.each do |part|
        zip.put_next_entry(zip_entry_for_part(part.sub(Regexp.new("#{tempdir}/?"), '')))
        zip.puts(File.read(part))
      end
    end

    def zip_entry_for_part(part)
      Zip::Entry.new("", part)
    end

    #
    # files
    #
    def parts
      Dir.glob(File.join(tempdir, "**", "*"), File::FNM_DOTMATCH).select {|f| File.file?(f)}
    end

    #
    # Prepare all of the format properties prior to passing them to Styles.rb.
    #
    def prepare_format_properties #:nodoc:
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
    def prepare_formats #:nodoc:
      @formats.formats.each do |format|
        xf_index  = format.xf_index
        dxf_index = format.dxf_index

        @xf_formats[xf_index] = format   if xf_index
        @dxf_formats[dxf_index] = format if dxf_index
      end
    end

    #
    # Iterate through the XF Format objects and give them an index to non-default
    # font elements.
    #
    def prepare_fonts #:nodoc:
      fonts = {}

      @xf_formats.each { |format| format.set_font_info(fonts) }

      @font_count = fonts.size

      # For the DXF formats we only need to check if the properties have changed.
      @dxf_formats.each do |format|
        # The only font properties that can change for a DXF format are: color,
        # bold, italic, underline and strikethrough.
        if format.color? || format.bold? || format.italic? || format.underline? || format.strikeout?
          format.has_dxf_font(true)
        end
      end
    end

    #
    # Iterate through the XF Format objects and give them an index to non-default
    # number format elements.
    #
    # User defined records start from index 0xA4.
    #
    def prepare_num_formats #:nodoc:
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
          num_format_count += 1 if ptrue?(format.xf_index)
        end
      end

      @num_format_count = num_format_count
    end

    #
    # Iterate through the XF Format objects and give them an index to non-default
    # border elements.
    #
    def prepare_borders #:nodoc:
      borders = {}

      @xf_formats.each { |format| format.set_border_info(borders) }

      @border_count = borders.size

      # For the DXF formats we only need to check if the properties have changed.
      @dxf_formats.each do |format|
        key = format.get_border_key
        format.has_dxf_border(true) if key =~ /[^0:]/
      end
    end

    #
    # Iterate through the XF Format objects and give them an index to non-default
    # fill elements.
    #
    # The user defined fill properties start from 2 since there are 2 default
    # fills: patternType="none" and patternType="gray125".
    #
    def prepare_fills #:nodoc:
      fills = {}
      index = 2    # Start from 2. See above.

      # Add the default fills.
      fills['0:0:0']  = 0
      fills['17:0:0'] = 1

      # Store the DXF colors separately since them may be reversed below.
      @dxf_formats.each do |format|
        if  format.pattern != 0 || format.bg_color != 0 || format.fg_color != 0
          format.has_dxf_fill(true)
          format.dxf_bg_color = format.bg_color
          format.dxf_fg_color = format.fg_color
        end
      end

      @xf_formats.each do |format|
        # The following logical statements jointly take care of special cases
        # in relation to cell colours and patterns:
        # 1. For a solid fill (_pattern == 1) Excel reverses the role of
        #    foreground and background colours, and
        # 2. If the user specifies a foreground or background colour without
        #    a pattern they probably wanted a solid fill, so we fill in the
        #    defaults.
        #
        if format.pattern == 1 && ne_0?(format.bg_color) && ne_0?(format.fg_color)
          format.fg_color, format.bg_color = format.bg_color, format.fg_color
        elsif format.pattern <= 1 && ne_0?(format.bg_color) && eq_0?(format.fg_color)
          format.fg_color = format.bg_color
          format.bg_color = 0
          format.pattern  = 1
        elsif format.pattern <= 1 && eq_0?(format.bg_color) && ne_0?(format.fg_color)
          format.bg_color = 0
          format.pattern  = 1
        end

        key = format.get_fill_key

        if fills[key]
          # Fill has already been used.
          format.fill_index = fills[key]
          format.has_fill(false)
        else
          # This is a new fill.
          fills[key]        = index
          format.fill_index = index
          format.has_fill(true)
          index += 1
        end
      end

      @fill_count = index
    end

    def eq_0?(val)
      ptrue?(val) ? false : true
    end

    def ne_0?(val)
      !eq_0?(val)
    end

    #
    # Iterate through the worksheets and store any defined names in addition to
    # any user defined names. Stores the defined names for the Workbook.xml and
    # the named ranges for App.xml.
    #
    def prepare_defined_names #:nodoc:
      @worksheets.each do |sheet|
        # Check for Print Area settings.
        if sheet.autofilter_area
          @defined_names << [
                             '_xlnm._FilterDatabase',
                             sheet.index,
                             sheet.autofilter_area,
                             1
                            ]
        end

        # Check for Print Area settings.
        if !sheet.print_area.empty?
          @defined_names << [
                             '_xlnm.Print_Area',
                             sheet.index,
                             sheet.print_area
                            ]
        end

        # Check for repeat rows/cols. aka, Print Titles.
        if !sheet.print_repeat_cols.empty? || !sheet.print_repeat_rows.empty?
          if !sheet.print_repeat_cols.empty? && !sheet.print_repeat_rows.empty?
            range = sheet.print_repeat_cols + ',' + sheet.print_repeat_rows
          else
            range = sheet.print_repeat_cols + sheet.print_repeat_rows
          end

          # Store the defined names.
          @defined_names << ['_xlnm.Print_Titles', sheet.index, range]
        end
      end

      @defined_names  = sort_defined_names(@defined_names)
      @named_ranges  = extract_named_ranges(@defined_names)
    end

    #
    # Iterate through the worksheets and set up the VML objects.
    #
    def prepare_vml_objects  #:nodoc:
      comment_id     = 0
      vml_drawing_id = 0
      vml_data_id    = 1
      vml_header_id  = 0
      vml_shape_id   = 1024
      comment_files  = 0
      has_button     = false

      @worksheets.each do |sheet|
        next if !sheet.has_vml? && !sheet.has_header_vml?
        if sheet.has_vml?
          if sheet.has_comments?
            comment_files += 1
            comment_id    += 1
          end
          vml_drawing_id += 1

          sheet.prepare_vml_objects(vml_data_id, vml_shape_id,
                                    vml_drawing_id, comment_id)

          # Each VML file should start with a shape id incremented by 1024.
          vml_data_id  +=    1 * ( 1 + sheet.num_comments_block )
          vml_shape_id += 1024 * ( 1 + sheet.num_comments_block )
        end

        if sheet.has_header_vml?
          vml_header_id  += 1
          vml_drawing_id += 1
          sheet.prepare_header_vml_objects(vml_header_id, vml_drawing_id)
        end

        # Set the sheet vba_codename if it has a button and the workbook
        # has a vbaProject binary.
        unless sheet.buttons_data.empty?
          has_button = true
          if @vba_project && !sheet.vba_codename
            sheet.set_vba_name
          end
        end
      end

      add_font_format_for_cell_comments if num_comment_files > 0

      # Set the workbook vba_codename if one of the sheets has a button and
      # the workbook has a vbaProject binary.
      if has_button && @vba_project && !@vba_codename
        set_vba_name
      end
    end

    #
    # Set the table ids for the worksheet tables.
    #
    def prepare_tables
      table_id = 0

      sheets.each do |sheet|
        table_id += sheet.prepare_tables(table_id + 1)
      end
    end

    def add_font_format_for_cell_comments
      format = Format.new(
                          @formats,
                          :font          => 'Tahoma',
                          :size          => 8,
                          :color_indexed => 81,
                          :font_only     => 1
                          )

      format.get_xf_index
      @formats.formats << format
    end

    #
    # Add "cached" data to charts to provide the numCache and strCache data for
    # series and title/axis ranges.
    #
    def add_chart_data #:nodoc:
      worksheets = {}
      seen_ranges = {}

      # Map worksheet names to worksheet objects.
      @worksheets.each { |worksheet| worksheets[worksheet.name] = worksheet }

      # Build an array of the worksheet charts including any combined charts.
      @charts.collect { |chart| [chart, chart.combined] }.flatten.compact.
        each do |chart|
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
          unless worksheets[sheetname]
            raise "Unknown worksheet reference '#{sheetname} in range '#{range}' passed to add_series()\n"
          end

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
        if token.kind_of?(Hash)
          string = @shared_strings.string(token[:sst_id])

          # Ignore rich strings for now. Deparse later if necessary.
          if string =~ %r!^<r>! && string =~ %r!</r>$!
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
    def sort_defined_names(names) #:nodoc:
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
    def normalise_defined_name(name) #:nodoc:
      name.sub(/^_xlnm./, '').downcase
    end

    # Used in the above sort routine to normalise the worksheet names for the
    # secondary sort. Removes leading quote and lowercases the strings.
    def normalise_sheet_name(name) #:nodoc:
      name.sub(/^'/, '').downcase
    end

    #
    # Extract the named ranges from the sorted list of defined names. These are
    # used in the App.xml file.
    #
    def extract_named_ranges(defined_names) #:nodoc:
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
            name = "#{sheet_name}!#{xlnm_type}"
          elsif index != -1
            name = "#{sheet_name}!#{name}"
          end

          named_ranges << name
        end
      end

      named_ranges
    end

    #
    # Iterate through the worksheets and set up any chart or image drawings.
    #
    def prepare_drawings #:nodoc:
      chart_ref_id = 0
      image_ref_id = 0
      drawing_id   = 0
      @worksheets.each do |sheet|
        chart_count = sheet.charts.size
        image_count = sheet.images.size
        shape_count = sheet.shapes.size
        header_image_count = sheet.header_images.size
        footer_image_count = sheet.footer_images.size
        has_drawing = false

        # Check that some image or drawing needs to be processed.
        next if chart_count + image_count + shape_count + header_image_count + footer_image_count == 0

        # Don't increase the drawing_id header/footer images.
        if chart_count + image_count + shape_count > 0
          drawing_id += 1
          has_drawing = true
        end

        # Prepare the worksheet charts.
        sheet.charts.each_with_index do |chart, index|
          chart_ref_id += 1
          sheet.prepare_chart(index, chart_ref_id, drawing_id)
        end

        # Prepare the worksheet images.
        sheet.images.each_with_index do |image, index|
          type, width, height, name, x_dpi, y_dpi = get_image_properties(image[2])
          image_ref_id += 1
          sheet.prepare_image(index, image_ref_id, drawing_id, width, height, name, type, x_dpi, y_dpi)
        end

        # Prepare the worksheet shapes.
        sheet.shapes.each_with_index do |shape, index|
          sheet.prepare_shape(index, drawing_id)
        end

        # Prepare the header images.
        header_image_count.times do |index|
          filename = sheet.header_images[index][0]
          position = sheet.header_images[index][1]

          type, width, height, name, x_dpi, y_dpi =
            get_image_properties(filename)

          image_ref_id += 1

          sheet.prepare_header_image(image_ref_id, width, height,
                                     name, type, position, x_dpi, y_dpi)
        end

        # Prepare the footer images.
        footer_image_count.times do |index|
          filename = sheet.footer_images[index][0]
          position = sheet.footer_images[index][1]

          type, width, height, name, x_dpi, y_dpi =
            get_image_properties(filename)

          image_ref_id += 1

          sheet.prepare_header_image(image_ref_id, width, height,
                                     name, type, position, x_dpi, y_dpi)
        end

        if has_drawing
          drawing = sheet.drawing
          @drawings << drawing
        end
      end

      # Sort the workbook charts references into the order that the were
      # written from the worksheets above.
      @charts = @charts.select { |chart| chart.id != -1 }.
        sort_by { |chart| chart.id }

      @drawing_count = drawing_id
    end

    #
    # Extract information from the image file such as dimension, type, filename,
    # and extension. Also keep track of previously seen images to optimise out
    # any duplicates.
    #
    def get_image_properties(filename)
      # Note the image_id, and previous_images mechanism isn't currently used.
      x_dpi = 96
      y_dpi = 96

      # Open the image file and import the data.
      data = File.binread(filename)
      if data.unpack('x A3')[0] == 'PNG'
        # Test for PNGs.
        type, width, height, x_dpi, y_dpi = process_png(data)
        @image_types[:png] = 1
      elsif data.unpack('n')[0] == 0xFFD8
        # Test for JPEG files.
        type, width, height, x_dpi, y_dpi = process_jpg(data, filename)
        @image_types[:jpeg] = 1
      elsif data.unpack('A2')[0] == 'BM'
        # Test for BMPs.
        type, width, height = process_bmp(data, filename)
        @image_types[:bmp] = 1
      else
        # TODO. Add Image::Size to support other types.
        raise "Unsupported image format for file: #{filename}\n"
      end

      @images << [filename, type]

      [type, width, height, File.basename(filename), x_dpi, y_dpi]
    end

    #
    # Extract width and height information from a PNG file.
    #
    def process_png(data)
      type   = 'png'
      width  = 0
      height = 0
      x_dpi  = 96
      y_dpi  = 96

      offset = 8
      data_length = data.size

      # Search through the image data to read the height and width in th the
      # IHDR element. Also read the DPI in the pHYs element.
      while offset < data_length

        length = data[offset + 0, 4].unpack("N")[0]
        png_type   = data[offset + 4, 4].unpack("A4")[0]

        case png_type
        when "IHDR"
          width  = data[offset +  8, 4].unpack("N")[0]
          height = data[offset + 12, 4].unpack("N")[0]
        when "pHYs"
          x_ppu = data[offset +  8, 4].unpack("N")[0]
          y_ppu = data[offset + 12, 4].unpack("N")[0]
          units = data[offset + 16, 1].unpack("C")[0]

          if units == 1
            x_dpi = x_ppu * 0.0254
            y_dpi = y_ppu * 0.0254
          end
        end

        offset = offset + length + 12

        break if png_type == "IEND"
      end
      raise "#{filename}: no size data found in png image.\n" unless height

      [type, width, height, x_dpi, y_dpi]
    end

    def process_jpg(data, filename)
      type     = 'jpeg'
      x_dpi    = 96
      y_dpi    = 96

      offset = 2
      data_length = data.bytesize

      # Search through the image data to read the height and width in the
      # 0xFFC0/C2 element. Also read the DPI in the 0xFFE0 element.
      while offset < data_length
        marker  = data[offset+0, 2].unpack("n")[0]
        length  = data[offset+2, 2].unpack("n")[0]

        if marker == 0xFFC0 || marker == 0xFFC2
          height = data[offset+5, 2].unpack("n")[0]
          width  = data[offset+7, 2].unpack("n")[0]
        end
        if marker == 0xFFE0
          units     = data[offset + 11, 1].unpack("C")[0]
          x_density = data[offset + 12, 2].unpack("n")[0]
          y_density = data[offset + 14, 2].unpack("n")[0]

          if units == 1
            x_dpi = x_density
            y_dpi = y_density
          elsif units == 2
            x_dpi = x_density * 2.54
            y_dpi = y_density * 2.54
          end
        end

        offset += length + 2
        break if marker == 0xFFDA
      end

      raise "#{filename}: no size data found in jpeg image.\n" unless height
      [type, width, height, x_dpi, y_dpi]
    end

    # Extract width and height information from a BMP file.
    def process_bmp(data, filename)       #:nodoc:
      type     = 'bmp'

      # Check that the file is big enough to be a bitmap.
      raise "#{filename} doesn't contain enough data." if data.bytesize <= 0x36

      # Read the bitmap width and height. Verify the sizes.
      width, height = data.unpack("x18 V2")
      raise "#{filename}: largest image width #{width} supported is 65k." if width > 0xFFFF
      raise "#{filename}: largest image height supported is 65k." if height > 0xFFFF

      # Read the bitmap planes and bpp data. Verify them.
      planes, bitcount = data.unpack("x26 v2")
      raise "#{filename} isn't a 24bit true color bitmap." unless bitcount == 24
      raise "#{filename}: only 1 plane supported in bitmap image." unless planes == 1

      # Read the bitmap compression. Verify compression.
      compression = data.unpack("x30 V")[0]
      raise "#{filename}: compression not supported in bitmap image." unless compression == 0
      [type, width, height]
    end
  end
end
