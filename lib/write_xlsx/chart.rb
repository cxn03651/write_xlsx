# -*- coding: utf-8 -*-
require 'write_xlsx/package/xml_writer_simple'
require 'write_xlsx/utility'

module Writexlsx
  # ==SYNOPSIS
  #
  # To create a simple Excel file with a chart using WriteXLSX:
  #
  #     require 'rubygems'
  #     require 'write_xlsx'
  #
  #     workbook  = WriteXLSX.new( 'chart.xlsx' )
  #     worksheet = workbook.add_worksheet
  #
  #     # Add the worksheet data the chart refers to.
  #     data = [
  #         [ 'Category', 2, 3, 4, 5, 6, 7 ],
  #         [ 'Value',    1, 4, 5, 2, 1, 5 ]
  #     ]
  #
  #     worksheet.write( 'A1', data )
  #
  #     # Add a worksheet chart.
  #     chart = workbook.add_chart( type => 'column' )
  #
  #     # Configure the chart.
  #     chart.add_series(
  #         :categories => '=Sheet1!$A$2:$A$7',
  #         :values     => '=Sheet1!$B$2:$B$7'
  #     )
  #
  #     workbook.close
  #
  # ==DESCRIPTION
  #
  # The Chart module is an abstract base class for modules that implement
  # charts in WriteXLSX. The information below is applicable to all of
  # the available subclasses.
  #
  # The Chart module isn't used directly. A chart object is created via
  # the Workbook add_chart() method where the chart type is specified:
  #
  #     chart = workbook.add_chart( :type => 'column' )
  #
  # Currently the supported chart types are:
  #
  # ===area
  # Creates an Area (filled line) style chart. See Writexlsx::Chart::Area.
  #
  # ===bar
  # Creates a Bar style (transposed histogram) chart. See Writexlsx::Chart::Bar.
  #
  # ===column
  # Creates a column style (histogram) chart. See Writexlsx::Chart::Column.
  #
  # ===line
  # Creates a Line style chart. See Writexlsx::Chart::Line.
  #
  # ===pie
  # Creates an Pie style chart. See Writexlsx::Chart::Pie.
  #
  # ===scatter
  # Creates an Scatter style chart. See Writexlsx::Chart::Scatter.
  #
  # ===stock
  # Creates an Stock style chart. See Writexlsx::Chart::Stock.
  #
  # ==CHART FORMATTING
  #
  # The following chart formatting properties can be set for any chart object
  # that they apply to (and that are supported by WriteXLSX) such
  # as chart lines, column fill areas, plot area borders, markers and other
  # chart elements documented above.
  #
  #     line
  #     border
  #     fill
  #     marker
  #     trendline
  #     data_labels
  # Chart formatting properties are generally set using hash refs.
  #
  #     chart.add_series(
  #         :values     => '=Sheet1!$B$1:$B$5',
  #         :line       => { color => 'blue' }
  #     )
  # In some cases the format properties can be nested. For example a marker
  # may contain border and fill sub-properties.
  #
  #     chart.add_series(
  #         :values     => '=Sheet1!$B$1:$B$5',
  #         :line       => { color => 'blue' },
  #         :marker     => {
  #             :type    => 'square',
  #             :size    => 5,
  #             :border  => { color => 'red' },
  #             :fill    => { color => 'yellow' }
  #         }
  #     )
  # ===Line
  #
  # The line format is used to specify properties of line objects that appear
  # in a chart such as a plotted line on a chart or a border.
  #
  # The following properties can be set for line formats in a chart.
  #
  #     none
  #     color
  #     width
  #     dash_type
  # The none property is uses to turn the line off (it is always on by default
  # except in Scatter charts). This is useful if you wish to plot a series
  # with markers but without a line.
  #
  #     chart.add_series(
  #         :values     => '=Sheet1!$B$1:$B$5',
  #         :line       => { none => 1 }
  #     )
  # The color property sets the color of the line.
  #
  #     chart.add_series(
  #         :values     => '=Sheet1!$B$1:$B$5',
  #         :line       => { color => 'red' }
  #     )
  # The available colors are shown in the main WriteXLSX documentation.
  # It is also possible to set the color of a line with a HTML style RGB color:
  #
  #     chart.add_series(
  #         :line       => { color => '#FF0000' }
  #     )
  # The width property sets the width of the line. It should be specified
  # in increments of 0.25 of a point as in Excel.
  #
  #     chart.add_series(
  #         :values     => '=Sheet1!$B$1:$B$5',
  #         :line       => { width => 3.25 }
  #     )
  # The dash_type property sets the dash style of the line.
  #
  #     chart->add_series(
  #         :values     => '=Sheet1!$B$1:$B$5',
  #         :line       => { dash_type => 'dash_dot' }
  #     )
  # The following dash_type values are available. They are shown in the
  # order that they appear in the Excel dialog.
  #
  #     solid
  #     round_dot
  #     square_dot
  #     dash
  #     dash_dot
  #     long_dash
  #     long_dash_dot
  #     long_dash_dot_dot
  # The default line style is solid.
  #
  # More than one line property can be specified at time:
  #
  #     chart.add_series(
  #         :values     => '=Sheet1!$B$1:$B$5',
  #         :line       => {
  #             :color     => 'red',
  #             :width     => 1.25,
  #             :dash_type => 'square_dot'
  #         }
  #     )
  # ===Border
  #
  # The border property is a synonym for line.
  #
  # It can be used as a descriptive substitute for line in chart types such
  # as Bar and Column that have a border and fill style rather than a line
  # style. In general chart objects with a border property will also have a
  # fill property.
  #
  # ===Fill
  #
  # The fill format is used to specify filled areas of chart objects such
  # as the interior of a column or the background of the chart itself.
  #
  # The following properties can be set for fill formats in a chart.
  #
  #     none
  #     color
  # The none property is uses to turn the fill property off (it is
  # generally on by default).
  #
  #     chart.add_series(
  #         :values     => '=Sheet1!$B$1:$B$5',
  #         :fill       => { none => 1 }
  #     )
  # The color property sets the color of the fill area.
  #
  #     chart.add_series(
  #         :values     => '=Sheet1!$B$1:$B$5',
  #         :fill       => { color => 'red' }
  #     )
  # The available colors are shown in the main WriteXLSX documentation.
  # It is also possible to set the color of a fill with a HTML style RGB color:
  #
  #     chart.add_series(
  #         :fill       => { color => '#FF0000' }
  #     )
  # The fill format is generally used in conjunction with a border format
  # which has the same properties as a line format.
  #
  #     chart.add_series(
  #         :values     => '=Sheet1!$B$1:$B$5',
  #         :border     => { color => 'red' },
  #         :fill       => { color => 'yellow' }
  #     )
  # ===Marker
  #
  # The marker format specifies the properties of the markers used to
  # distinguish series on a chart. In general only Line and Scatter
  # chart types and trendlines use markers.
  #
  # The following properties can be set for marker formats in a chart.
  #
  #     type
  #     size
  #     border
  #     fill
  # The type property sets the type of marker that is used with a series.
  #
  #     chart.add_series(
  #         :values     => '=Sheet1!$B$1:$B$5',
  #         :marker     => { type => 'diamond' }
  #     )
  # The following type properties can be set for marker formats in a chart.
  # These are shown in the same order as in the Excel format dialog.
  #
  #     automatic
  #     none
  #     square
  #     diamond
  #     triangle
  #     x
  #     star
  #     short_dash
  #     long_dash
  #     circle
  #     plus
  # The automatic type is a special case which turns on a marker using the
  # default marker style for the particular series number.
  #
  #     chart.add_series(
  #         :values     => '=Sheet1!$B$1:$B$5',
  #         :marker     => { type => 'automatic' }
  #     )
  # If automatic is on then other marker properties such as size,
  # border or fill cannot be set.
  #
  # The size property sets the size of the marker and is generally used in
  # conjunction with type.
  #
  #     chart.add_series(
  #         :values     => '=Sheet1!$B$1:$B$5',
  #         :marker     => { type => 'diamond', size => 7 }
  #     )
  # Nested border and fill properties can also be set for a marker.
  # These have the same sub-properties as shown above.
  #
  #     chart.add_series(
  #         :values     => '=Sheet1!$B$1:$B$5',
  #         :marker     => {
  #             :type    => 'square',
  #             :size    => 5,
  #             :border  => { color => 'red' },
  #             :fill    => { color => 'yellow' }
  #         }
  #     )
  # ===Trendline
  #
  # A trendline can be added to a chart series to indicate trends in the data
  # such as a moving average or a polynomial fit.
  #
  # The following properties can be set for trendline formats in a chart.
  #
  #     type
  #     order       (for polynomial trends)
  #     period      (for moving average)
  #     forward     (for all except moving average)
  #     backward    (for all except moving average)
  #     name
  #     line
  # The type property sets the type of trendline in the series.
  #
  #     chart.add_series(
  #         :values     => '=Sheet1!$B$1:$B$5',
  #         :trendline  => { type => 'linear' }
  #     )
  # The available trendline types are:
  #
  #     exponential
  #     linear
  #     log
  #     moving_average
  #     polynomial
  #     power
  # A polynomial trendline can also specify the order of the polynomial.
  # The default value is 2.
  #
  #     chart.add_series(
  #         :values    => '=Sheet1!$B$1:$B$5',
  #         :trendline => {
  #             :type  => 'polynomial',
  #             :order => 3
  #         }
  #     )
  # A moving_average trendline can also the period of the moving average.
  # The default value is 2.
  #
  #     chart.add_series(
  #         :values     => '=Sheet1!$B$1:$B$5',
  #         :trendline  => {
  #             :type   => 'moving_average',
  #             :period => 3
  #         }
  #     )
  # The forward and backward properties set the forecast period of the
  # trendline.
  #
  #     chart.add_series(
  #         :values    => '=Sheet1!$B$1:$B$5',
  #         :trendline => {
  #             :type     => 'linear',
  #             :forward  => 0.5,
  #             :backward => 0.5
  #         }
  #     )
  # The name property sets an optional name for the trendline that will
  # appear in the chart legend. If it isn't specified the Excel default
  # name will be displayed. This is usually a combination of the trendline
  # type and the series name.
  #
  #     chart.add_series(
  #         :values    => '=Sheet1!$B$1:$B$5',
  #         :trendline => {
  #             :type => 'linear',
  #             :name => 'Interpolated trend'
  #         }
  #     )
  # Several of these properties can be set in one go:
  #
  #     chart.add_series(
  #         :values     => '=Sheet1!$B$1:$B$5',
  #         :trendline  => {
  #             :type     => 'linear',
  #             :name     => 'My trend name',
  #             :forward  => 0.5,
  #             :backward => 0.5,
  #             :line     => {
  #                 :color     => 'red',
  #                 :width     => 1,
  #                 :dash_type => 'long_dash'
  #             }
  #         }
  #     )
  # Trendlines cannot be added to series in a stacked chart or pie chart or
  # (when implemented) to 3-D, radar, surface, or doughnut charts.
  #
  # ==Data Labels
  #
  # Data labels can be added to a chart series to indicate the values of
  # the plotted data points.
  #
  # The following properties can be set for data_labels formats in a chart.
  #
  #     value
  #     category
  #     series_name
  # The value property turns on the Value data label for a series.
  #
  #     chart.add_series(
  #         :values      => '=Sheet1!$B$1:$B$5',
  #         :data_labels => { value => 1 }
  #     )
  # The category property turns on the Category Name data label for a series.
  #
  #     chart.add_series(
  #         :values      => '=Sheet1!$B$1:$B$5',
  #         :data_labels => { category => 1 }
  #     )
  # The series_name property turns on the Series Name data label for a series.
  #
  #     chart.add_series(
  #         :values      => '=Sheet1!$B$1:$B$5',
  #         :data_labels => { series_name => 1 }
  #     )
  # Other formatting options
  #
  # Other formatting options will be added in time. If there is a feature that
  # you would like to see included drop me a line.
  #
  class Chart
    include Writexlsx::Utility

    attr_accessor :id   # :nodoc:
    attr_writer :index, :palette   # :nodoc:
    attr_reader :embedded, :formula_ids, :formula_data   # :nodoc:

    #
    # Factory method for returning chart objects based on their class type.
    #
    def self.factory(chart_subclass) # :nodoc:
      case chart_subclass.downcase.capitalize
      when 'Area'
        require 'write_xlsx/chart/area'
        Chart::Area.new
      when 'Bar'
        require 'write_xlsx/chart/bar'
        Chart::Bar.new
      when 'Column'
        require 'write_xlsx/chart/column'
        Chart::Column.new
      when 'Line'
        require 'write_xlsx/chart/line'
        Chart::Line.new
      when 'Pie'
        require 'write_xlsx/chart/pie'
        Chart::Pie.new
      when 'Scatter'
        require 'write_xlsx/chart/scatter'
        Chart::Scatter.new
      when 'Stock'
        require 'write_xlsx/chart/stock'
        Chart::Stock.new
      end
    end

    def initialize(subtype)   # :nodoc:
      @writer = Package::XMLWriterSimple.new

      @subtype           = subtype
      @sheet_type        = 0x0200
      @orientation       = 0x0
      @series            = []
      @embedded          = 0
      @id                = ''
      @style_id          = 2
      @axis_ids          = []
      @has_category      = 0
      @requires_category = 0
      @legend_position   = 'right'
      @cat_axis_position = 'b'
      @val_axis_position = 'l'
      @formula_ids       = {}
      @formula_data      = []
      @horiz_cat_axis    = 0
      @horiz_val_axis    = 1
      @protection        = 0
      @x_axis            = {}
      @y_axis            = {}

      set_default_properties
    end

    def set_xml_writer(filename)   # :nodoc:
      @writer.set_xml_writer(filename)
    end

    #
    # Assemble and write the XML file.
    #
    def assemble_xml_file   # :nodoc:
      @writer.xml_decl

      # Write the c:chartSpace element.
      write_chart_space

      # Write the c:lang element.
      write_lang

      # Write the c:style element.
      write_style

      # Write the c:protection element.
      write_protection

      # Write the c:chart element.
      write_chart

      # Write the c:printSettings element.
      write_print_settings if @embedded

      # Close the worksheet tag.
      @writer.end_tag( 'c:chartSpace')

      # Close the XML writer object and filehandle.
      @writer.crlf
      @writer.close
    end

    #
    # Add a series and it's properties to a chart.
    #
    # In an Excel chart a "series" is a collection of information such as
    # values, x-axis labels and the formatting that define which data is
    # plotted.
    #
    # With a WriteXLSX chart object the add_series() method is used to
    # set the properties for a series:
    #
    #     chart.add_series(
    #         :categories => '=Sheet1!$A$2:$A$10', # Optional.
    #         :values     => '=Sheet1!$B$2:$B$10', # Required.
    #         :line       => { color => 'blue' }
    #     )
    #
    # The properties that can be set are:
    #
    # ====:values
    # This is the most important property of a series and must be set
    # for every chart object. It links the chart with the worksheet data
    # that it displays. A formula or array ref can be used for the
    # data range, see below.
    #
    # ====:categories
    # This sets the chart category labels. The category is more or less
    # the same as the X-axis. In most chart types the categories property
    # is optional and the chart will just assume a sequential series
    # from 1 .. n.
    #
    # ====:name
    # Set the name for the series. The name is displayed in the chart
    # legend and in the formula bar. The name property is optional and
    # if it isn't supplied it will default to Series 1 .. n.
    #
    # ====:line
    # Set the properties of the series line type such as colour and
    # width. See the "CHART FORMATTING" section below.
    #
    # ====:border
    # Set the border properties of the series such as colour and style.
    # See the "CHART FORMATTING" section below.
    #
    # ====:fill
    # Set the fill properties of the series such as colour. See the
    # "CHART FORMATTING" section below.
    #
    # ====:marker
    # Set the properties of the series marker such as style and color.
    # See the "CHART FORMATTING" section below.
    #
    # ====:trendline
    # Set the properties of the series trendline such as linear,
    # polynomial and moving average types. See the "CHART FORMATTING"
    # section below.
    #
    # ====:data_labels
    # Set data labels for the series. See the "CHART FORMATTING"
    # section below.
    #
    # ====:invert_if_negative
    # Invert the fill colour for negative values. Usually only applicable
    # to column and bar charts.
    #
    # The categories and values can take either a range formula such
    # as =Sheet1!$A$2:$A$7 or, more usefully when generating the range
    # programmatically, an array ref with zero indexed row/column values:
    #
    #      [ sheetname, row_start, row_end, col_start, col_end ]
    # The following are equivalent:
    #
    #     chart.add_series( categories => '=Sheet1!$A$2:$A$7'      ) # Same as ...
    #     chart.add_series( categories => [ 'Sheet1', 1, 6, 0, 0 ] ) # Zero-indexed.
    #
    # You can add more than one series to a chart. In fact, some chart
    # types such as stock require it. The series numbering and order in
    # the Excel chart will be the same as the order in which that are added
    # in WriteXLSX.
    #
    #     # Add the first series.
    #     chart.add_series(
    #         :categories => '=Sheet1!$A$2:$A$7',
    #         :values     => '=Sheet1!$B$2:$B$7',
    #         :name       => 'Test data series 1'
    #     )
    #
    #     # Add another series. Same categories. Different range values.
    #     chart.add_series(
    #         :categories => '=Sheet1!$A$2:$A$7',
    #         :values     => '=Sheet1!$C$2:$C$7',
    #         :name       => 'Test data series 2'
    #     )
    #
    def add_series(params)
      # Check that the required input has been specified.
      unless params.has_key?(:values)
        raise "Must specify ':values' in add_series"
      end

      if @requires_category != 0 && !params.has_key?(:categories)
        raise  "Must specify ':categories' in add_series for this chart type"
      end

      # Convert aref params into a formula string.
      values     = aref_to_formula(params[:values])
      categories = aref_to_formula(params[:categories])

      # Switch name and name_formula parameters if required.
      name, name_formula = process_names(params[:name], params[:name_formula])

      # Get an id for the data equivalent to the range formula.
      cat_id  = get_data_id(categories,   params[:categories_data])
      val_id  = get_data_id(values,       params[:values_data])
      name_id = get_data_id(name_formula, params[:name_data])

      # Set the line properties for the series.
      line = get_line_properties(params[:line])

      # Allow 'border' as a synonym for 'line' in bar/column style charts.
      line = get_line_properties(params[:border]) if params[:border]

      # Set the fill properties for the series.
      fill = get_fill_properties(params[:fill])

      # Set the marker properties for the series.
      marker = get_marker_properties(params[:marker])

      # Set the trendline properties for the series.
      trendline = get_trendline_properties(params[:trendline])

      # Set the labels properties for the series.
      labels = get_labels_properties(params[:data_labels])

      # Set the "invert if negative" fill property.
      invert_if_neg = params[:invert_if_negative]

      # Add the user supplied data to the internal structures.
      @series << {
        :_values        => values,
        :_categories    => categories,
        :_name          => name,
        :_name_formula  => name_formula,
        :_name_id       => name_id,
        :_val_data_id   => val_id,
        :_cat_data_id   => cat_id,
        :_line          => line,
        :_fill          => fill,
        :_marker        => marker,
        :_trendline     => trendline,
        :_labels        => labels,
        :_invert_if_neg => invert_if_neg
      }
    end

    #
    # Set the properties of the X-axis.
    #
    # The set_x_axis() method is used to set properties of the X axis.
    #
    #     chart.set_x_axis( :name => 'Quarterly results' )
    #
    # The properties that can be set are:
    #
    #     :name
    #     :min
    #     :max
    #     :minor_unit
    #     :major_unit
    #     :crossing
    #     :reverse
    #     :log_base
    #     :label_position
    #
    # These are explained below. Some properties are only applicable to value
    # or category axes, as indicated. See "Value and Category Axes" for an
    # explanation of Excel's distinction between the axis types.
    #
    # ====:name
    # Set the name (title or caption) for the axis. The name is displayed
    # below the X axis. The name property is optional. The default is to
    # have no axis name. (Applicable to category and value axes).
    #
    #     chart.set_x_axis( :name => 'Quarterly results' )
    #
    # The name can also be a formula such as =Sheet1!$A$1.
    #
    # ====:min
    # Set the minimum value for the axis range.
    # (Applicable to value axes only).
    #
    #     chart.set_x_axis( :min => 20 )
    # ====:max
    # Set the maximum value for the axis range.
    # (Applicable to value axes only).
    #
    #     chart.set_x_axis( :max => 80 )
    # ====:minor_unit
    # Set the increment of the minor units in the axis range.
    # (Applicable to value axes only).
    #
    #     chart.set_x_axis( :minor_unit => 0.4 )
    # ====:major_unit
    # Set the increment of the major units in the axis range.
    # (Applicable to value axes only).
    #
    #     chart.set_x_axis( :major_unit => 2 )
    # ====:crossing
    # Set the position where the y axis will cross the x axis.
    # (Applicable to category and value axes).
    #
    # The crossing value can either be the string 'max' to set the crossing
    # at the maximum axis value or a numeric value.
    #
    #     chart.set_x_axis( :crossing => 3 )
    #     # or
    #     chart.set_x_axis( :crossing => 'max' )
    # For category axes the numeric value must be an integer to represent
    # the category number that the axis crosses at. For value axes it can
    # have any value associated with the axis.
    #
    # If crossing is omitted (the default) the crossing will be set
    # automatically by Excel based on the chart data.
    #
    # ====:reverse
    # Reverse the order of the axis categories or values.
    # (Applicable to category and value axes).
    #
    #     chart.set_x_axis( :reverse => 1 )
    # ====:log_base
    # Set the log base of the axis range.
    # (Applicable to value axes only).
    #
    #     chart.set_x_axis( :log_base => 10 )
    # ====:label_position
    # Set the "Axis labels" position for the axis.
    # The following positions are available:
    #
    #     next_to (the default)
    #     high
    #     low
    #     none
    # More than one property can be set in a call to set_x_axis:
    #
    #     chart.set_x_axis(
    #         :name => 'Quarterly results',
    #         :min  => 10,
    #         :max  => 80
    #     )
    #
    def set_x_axis(params)
      @x_axis = convert_axis_args(params)
    end

    #
    # Set the properties of the Y-axis.
    #
    # The set_y_axis() method is used to set properties of the Y axis.
    # The properties that can be set are the same as for set_x_axis,
    #
    def set_y_axis(params)
      @y_axis = convert_axis_args(params)
    end

    #
    # Set the properties of the chart title.
    #
    # The set_title() method is used to set properties of the chart title.
    #
    #     chart.set_title( :name => 'Year End Results' )
    # The properties that can be set are:
    #
    # ====:name
    # Set the name (title) for the chart. The name is displayed above the
    # chart. The name can also be a formula such as =Sheet1!$A$1. The name
    # property is optional. The default is to have no chart title.
    #
    def set_title(params)
      name, name_formula = process_names(params[:name], params[:name_formula])

      data_id = get_data_id(name_formula, params[:data])

      @title_name    = name
      @title_formula = name_formula
      @title_data_id = data_id
    end

    #
    # Set the properties of the chart legend.
    #
    # The set_legend() method is used to set properties of the chart legend.
    #
    #     chart.set_legend( :position => 'none' )
    # The properties that can be set are:
    #
    # ====:position
    # Set the position of the chart legend.
    #
    #     chart.set_legend( :position => 'bottom' )
    # The default legend position is right. The available positions are:
    #
    #     none
    #     top
    #     bottom
    #     left
    #     right
    #     overlay_left
    #     overlay_right
    # ====:delete_series
    # This allows you to remove 1 or more series from the the legend
    # (the series will still display on the chart). This property takes
    # an array ref as an argument and the series are zero indexed:
    #
    #     # Delete/hide series index 0 and 2 from the legend.
    #     chart.set_legend( :delete_series => [0, 2] )
    #
    def set_legend(params)
      @legend_position = params[:position] || 'right'
      @legend_delete_series = params[:delete_series]
    end

    #
    # Set the properties of the chart plotarea.
    #
    # The set_plotarea() method is used to set properties of the plot area
    # of a chart.
    #
    # This method isn't implemented yet and is only available in
    # writeexcel gem. However, it can be simulated using the
    # set_style() method.
    #
    def set_plotarea(params)

      # TODO. Need to refactor for XLSX format.
      return

      return if params.empty?

      area = @plotarea

      # Set the plotarea visibility.
      if params[:visible]
        area[:_visible] = params[:visible]
        return unless area[:_visible]
      end

      # TODO. could move this out of if statement.
      area[:_bg_color_index] = 0x08

      # Set the chart background colour.
      if params[:color]
        index, rgb = get_color_indices(params[:color])
        if index
          area[:_fg_color_index] = index
          area[:_fg_color_rgb]   = rgb
          area[:_bg_color_index] = 0x08
          area[:_bg_color_rgb]   = 0x000000
        end

      end

      # Set the border line colour.
      if params[:line_color]
        index, rgb = get_color_indices(params[:line_color])
        if index
          area[:_line_color_index] = index
          area[:_line_color_rgb]   = rgb
        end
      end

      # Set the border line pattern.
      if params[:line_pattern]
        pattern = get_line_pattern(params[:line_pattern])
        area[:_line_pattern] = pattern
      end

      # Set the border line weight.
      if params[:line_weight]
        weight = get_line_weight(params[:line_weight])
        area[:_line_weight] = weight
      end
    end

    #
    # Set the properties of the chart chartarea.
    #
    # The set_chartarea() method is used to set the properties of the chart
    # area.
    #
    # This method isn't implemented yet and is only available in
    # writeexcel gem. However, it can be simulated using the
    # set_style() method.
    #
    def set_chartarea(params)
      # TODO. Need to refactor for XLSX format.
      return

      return if params.empty?

      area = @chartarea

      # Embedded automatic line weight has a different default value.
      area[:_line_weight] = 0xFFFF if @embedded

      # Set the chart background colour.
      if params[:color]
        index, rgb = get_color_indices(params[:color])
        if index
          area[:_fg_color_index] = index
          area[:_fg_color_rgb]   = rgb
          area[:_bg_color_index] = 0x08
          area[:_bg_color_rgb]   = 0x000000
          area[:_area_pattern]   = 1
          area[:_area_options]   = 0x0000 if @embedded
          area[:_visible]        = 1
        end
      end

      # Set the border line colour.
      if params[:line_color]
        index, rgb = get_color_indices(params[:line_color])
        if index
          area[:_line_color_index] = index
          area[:_line_color_rgb]   = rgb
          area[:_line_pattern]     = 0x00
          area[:_line_options]     = 0x0000
          area[:_visible]          = 1
        end
      end

      # Set the border line pattern.
      if params[:line_pattern]
        pattern = get_line_pattern(params[:line_pattern])
        area[:_line_pattern]     = pattern
        area[:_line_options]     = 0x0000
        area[:_line_color_index] = 0x4F unless params[:line_color]
        area[:_visible]          = 1
      end

      # Set the border line weight.
      if params[:line_weight]
        weight = get_line_weight(params[:line_weight])
        area[:_line_weight]      = weight
        area[:_line_options]     = 0x0000
        area[:_line_pattern]     = 0x00 unless params[:line_pattern]
        area[:_line_color_index] = 0x4F unless params[:line_color]
        area[:_visible]          = 1
      end
    end

    #
    # Set on of the 42 built-in Excel chart styles. The default style is 2.
    #
    # The set_style() method is used to set the style of the chart to one
    # of the 42 built-in styles available on the 'Design' tab in Excel:
    #
    #     chart.set_style( 4 )
    #
    def set_style(style_id = 2)
      style_id = 2 if style_id < 0 || style_id > 42
      @style_id = style_id
    end

    #
    # Setup the default configuration data for an embedded chart.
    #
    def set_embedded_config_data
      @embedded = 1

      # TODO. We may be able to remove this after refactoring.

      @chartarea = {
        :_visible          => 1,
        :_fg_color_index   => 0x4E,
        :_fg_color_rgb     => 0xFFFFFF,
        :_bg_color_index   => 0x4D,
        :_bg_color_rgb     => 0x000000,
        :_area_pattern     => 0x0001,
        :_area_options     => 0x0001,
        :_line_pattern     => 0x0000,
        :_line_weight      => 0x0000,
        :_line_color_index => 0x4D,
        :_line_color_rgb   => 0x000000,
        :_line_options     => 0x0009
      }

    end

      #
      # Write the <c:barChart> element.
      #
      def write_bar_chart   # :nodoc:
        subtype = @subtype
        subtype = 'percentStacked' if subtype == 'percent_stacked'

        @writer.tag_elements('c:barChart') do
          # Write the c:barDir element.
          write_bar_dir
          # Write the c:grouping element.
          write_grouping(subtype)
          # Write the series elements.
          write_series
        end
      end

    private

    #
    # Convert user defined axis values into private hash values.
    #
    def convert_axis_args(params) # :nodoc:
      name, name_formula = process_names(params[:name], params[:name_formula])

      data_id = get_data_id(name_formula, params[:data])

      axis = {
        :_name            => name,
        :_formula         => name_formula,
        :_data_id         => data_id,
        :_reverse         => params[:reverse],
        :_min             => params[:min],
        :_max             => params[:max],
        :_minor_unit      => params[:minor_unit],
        :_major_unit      => params[:major_unit],
        :_minor_unit_type => params[:minor_unit_type],
        :_major_unit_type => params[:major_unit_type],
        :_log_base        => params[:log_base],
        :_crossing        => params[:crossing],
        :_position        => params[:position],
        :_label_position  => params[:label_position]
      }

      # Only use the first letter of bottom, top, left or right.
      axis[:_position] = axis[:_position].downcase[0, 1] if axis[:_position]

      axis
    end

    #
    # Convert and aref of row col values to a range formula.
    #
    def aref_to_formula(data) # :nodoc:
      # If it isn't an array ref it is probably a formula already.
      return data unless data.kind_of?(Array)
      xl_range_formula(*data)
    end

    #
    # Switch name and name_formula parameters if required.
    #
    def process_names(name = nil, name_formula = nil) # :nodoc:
      # Name looks like a formula, use it to set name_formula.
      if name && name =~ /^=[^!]+!\$/
        name_formula = name
        name         = ''
      end

      [name, name_formula]
    end

    #
    # Find the overall type of the data associated with a series.
    #
    # TODO. Need to handle date type.
    #
    def get_data_type(data) # :nodoc:
      # Check for no data in the series.
      return 'none' unless data
      return 'none' if data.empty?

      # If the token isn't a number assume it is a string.
      data.each do |token|
        next unless token
        return 'str' unless token.kind_of?(Numeric)
      end

      # The series data was all numeric.
      'num'
    end

    #
    # Assign an id to a each unique series formula or title/axis formula. Repeated
    # formulas such as for categories get the same id. If the series or title
    # has user specified data associated with it then that is also stored. This
    # data is used to populate cached Excel data when creating a chart.
    # If there is no user defined data then it will be populated by the parent
    # workbook in Workbook::_add_chart_data
    #
    def get_data_id(formula, data) # :nodoc:
      # Ignore series without a range formula.
      return unless formula

      # Strip the leading '=' from the formula.
      formula = formula.sub(/^=/, '')

      # Store the data id in a hash keyed by the formula and store the data
      # in a separate array with the same id.
      if !@formula_ids.has_key?(formula)
        # Haven't seen this formula before.
        id = @formula_data.size

        @formula_data << data
        @formula_ids[formula] = id
      else
        # Formula already seen. Return existing id.
        id = @formula_ids[formula]

        # Store user defined data if it isn't already there.
        @formula_data[id] = data unless @formula_data[id]
      end

      id
    end


    #
    # Convert the user specified colour index or string to a rgb colour.
    #
    def get_color(color) # :nodoc:
      # Convert a HTML style #RRGGBB color.
      if color and color =~ /^#[0-9a-fA-F]{6}$/
        color = color.sub(/^#/, '')
        return color.upperca
      end

      index = Format.get_color(color)

      # Set undefined colors to black.
      unless index
        index = 0x08
        raise "Unknown color '#{color}' used in chart formatting."
      end

      get_palette_color(index)
    end

    #
    # Convert from an Excel internal colour index to a XML style #RRGGBB index
    # based on the default or user defined values in the Workbook palette.
    # Note: This version doesn't add an alpha channel.
    #
    def get_palette_color(index) # :nodoc:
      palette = @palette

      # Adjust the colour index.
      index -= 8

      # Palette is passed in from the Workbook class.
      rgb = palette[index]

      sprintf("%02X%02X%02X", *rgb)
    end

    #
    # Get the Excel chart index for line pattern that corresponds to the user
    # defined value.
    #
    def get_line_pattern(value) # :nodoc:
      value = value.downcase
      default = 0

      patterns = {
        0              => 5,
        1              => 0,
        2              => 1,
        3              => 2,
        4              => 3,
        5              => 4,
        6              => 7,
        7              => 6,
        8              => 8,
        'solid'        => 0,
        'dash'         => 1,
        'dot'          => 2,
        'dash-dot'     => 3,
        'dash-dot-dot' => 4,
        'none'         => 5,
        'dark-gray'    => 6,
        'medium-gray'  => 7,
        'light-gray'   => 8
      }

      if patterns.has_key(:value)
        pattern = patterns[:value]
      else
        pattern = default
      end

      pattern
    end

    #
    # Get the Excel chart index for line weight that corresponds to the user
    # defined value.
    #
    def get_line_weight(value) # :nodoc:
      value = value.downcase
      default = 0

      weights = {
        1          => -1,
        2          => 0,
        3          => 1,
        4          => 2,
        'hairline' => -1,
        'narrow'   => 0,
        'medium'   => 1,
        'wide'     => 2
      }

      if weights[:value]
        weight = weights[:value]
      else
        weight = default
      end

      weight
    end

    #
    # Convert user defined line properties to the structure required internally.
    #
    def get_line_properties(line) # :nodoc:
      return { :_defined => 0 } unless line

      dash_types = {
        :solid               => 'solid',
        :round_dot           => 'sysDot',
        :square_dot          => 'sysDash',
        :dash                => 'dash',
        :dash_dot            => 'dashDot',
        :long_dash           => 'lgDash',
        :long_dash_dot       => 'lgDashDot',
        :long_dash_dot_dot   => 'lgDashDotDot',
        :dot                 => 'dot',
        :system_dash_dot     => 'sysDashDot',
        :system_dash_dot_dot => 'sysDashDotDot'
      }

      # Check the dash type.
      dash_type = line[:dash_type]

      if dash_type
        if dash_types[dash_type]
          line[dash_type] = dash_types[dash_type]
        else
          raise "Unknown dash type '#{dash_type}'\n"
        end
      end

      line[:_defined] = 1

      line
    end

    #
    # Convert user defined fill properties to the structure required internally.
    #
    def get_fill_properties(fill) # :nodoc:
      return { :_defined => 0 } unless fill

      fill[:_defined] = 1

      fill
    end

    #
    # Convert user defined marker properties to the structure required internally.
    #
    def get_marker_properties(marker) # :nodoc:
      return unless marker

      types = {
        :automatic  => 'automatic',
        :none       => 'none',
        :square     => 'square',
        :diamond    => 'diamond',
        :triangle   => 'triangle',
        :x          => 'x',
        :star       => 'start',
        :dot        => 'dot',
        :short_dash => 'dot',
        :dash       => 'dash',
        :long_dash  => 'dash',
        :circle     => 'circle',
        :plus       => 'plus',
        :picture    => 'picture'
      }

      # Check for valid types.
      marker_type = marker[type]

      if marker_type
        marker[automatic] = 1 if marker_type == 'automatic'

        if types[marker_type]
          marker[type] = types[marker_type]
        else
          raise "Unknown marker type '#{marker_type}'\n"
        end
      end

      # Set the line properties for the marker..
      line = get_line_properties(marker[line])

      # Allow 'border' as a synonym for 'line'.
      line = get_line_properties(marker[border]) if marker[border]

      # Set the fill properties for the marker.
      fill = get_fill_properties(marker[fill])

      marker[:_line] = line
      marker[:_fill] = fill

      marker
    end

    #
    # Convert user defined trendline properties to the structure required internally.
    #
    def get_trendline_properties(trendline) # :nodoc:
      return unless trendline

      types = {
        :exponential    => 'exp',
        :linear         => 'linear',
        :log            => 'log',
        :moving_average => 'movingAvg',
        :polynomial     => 'poly',
        :power          => 'power'
      }

      # Check the trendline type.
      trend_type = trendline[type]

      if types[trend_type]
        trendline[type] = types[trend_type]
      else
        raise "Unknown trendline type '#{trend_type}'\n"
      end

      # Set the line properties for the trendline..
      line = get_line_properties(trendline[line])

      # Allow 'border' as a synonym for 'line'.
      line = get_line_properties(trendline[border]) if trendline[border]

      # Set the fill properties for the trendline.
      fill = get_fill_properties(trendline[fill])

      trendline[:_line] = line
      trendline[:_fill] = fill

      return trendline
    end

    #
    # Convert user defined labels properties to the structure required internally.
    #
    def get_labels_properties(labels) # :nodoc:
      return nil unless labels

      return labels
    end

    #
    # Add a unique id for an axis.
    #
    def add_axis_id # :nodoc:
      chart_id   = 1 + @id
      axis_count = 1 + @axis_ids.size

      axis_id = sprintf('5%03d%04d', chart_id, axis_count)

      @axis_ids << axis_id

      axis_id
    end

    #
    # Setup the default properties for a chart.
    #
    def set_default_properties # :nodoc:
      @chartarea = {
        :_visible          => 0,
        :_fg_color_index   => 0x4E,
        :_fg_color_rgb     => 0xFFFFFF,
        :_bg_color_index   => 0x4D,
        :_bg_color_rgb     => 0x000000,
        :_area_pattern     => 0x0000,
        :_area_options     => 0x0000,
        :_line_pattern     => 0x0005,
        :_line_weight      => 0xFFFF,
        :_line_color_index => 0x4D,
        :_line_color_rgb   => 0x000000,
        :_line_options     => 0x0008
      }

      @plotarea = {
        :_visible          => 1,
        :_fg_color_index   => 0x16,
        :_fg_color_rgb     => 0xC0C0C0,
        :_bg_color_index   => 0x4F,
        :_bg_color_rgb     => 0x000000,
        :_area_pattern     => 0x0001,
        :_area_options     => 0x0000,
        :_line_pattern     => 0x0000,
        :_line_weight      => 0x0000,
        :_line_color_index => 0x17,
        :_line_color_rgb   => 0x808080,
        :_line_options     => 0x0000
      }
    end

    #
    # Write the <c:chartSpace> element.
    #
    def write_chart_space # :nodoc:
      schema  = 'http://schemas.openxmlformats.org/'
      xmlns_c = schema + 'drawingml/2006/chart'
      xmlns_a = schema + 'drawingml/2006/main'
      xmlns_r = schema + 'officeDocument/2006/relationships'

      attributes = [
                    'xmlns:c', xmlns_c,
                    'xmlns:a', xmlns_a,
                    'xmlns:r', xmlns_r
                   ]

      @writer.start_tag('c:chartSpace', attributes)
    end

    #
    # Write the <c:lang> element.
    #
    def write_lang # :nodoc:
      val  = 'en-US'

      attributes = ['val', val]

      @writer.empty_tag('c:lang', attributes)
    end

    #
    # Write the <c:style> element.
    #
    def write_style # :nodoc:
      style_id = @style_id

      # Don't write an element for the default style, 2.
      return if style_id == 2

      attributes = ['val', style_id]

      @writer.empty_tag('c:style', attributes)
    end

    #
    # Write the <c:chart> element.
    #
    def write_chart # :nodoc:
      @writer.tag_elements('c:chart') do
        # Write the chart title elements.
        if title = @title_formula
          write_title_formula(title, @title_data_id)
        elsif title = @title_name
          write_title_rich(title)
        end

        # Write the c:plotArea element.
        write_plot_area
        # Write the c:legend element.
        write_legend
        # Write the c:plotVisOnly element.
        write_plot_vis_only
      end
    end

    #
    # Write the <c:plotArea> element.
    #
    def write_plot_area # :nodoc:
      @writer.tag_elements('c:plotArea') do
        # Write the c:layout element.
        write_layout
        # Write the subclass chart type element.
        write_chart_type
        # Write the c:catAx element.
        write_cat_axis
        # Write the c:catAx element.
        write_val_axis
      end
    end

    #
    # Write the <c:layout> element.
    #
    def write_layout # :nodoc:
      @writer.empty_tag('c:layout')
    end

    #
    # Write the chart type element. This method should be overridden by the
    # subclasses.
    #
    def write_chart_type # :nodoc:
    end

    #
    # Write the <c:grouping> element.
    #
    def write_grouping(val) # :nodoc:
      attributes = ['val', val]
      @writer.empty_tag('c:grouping', attributes)
    end

    #
    # Write the series elements.
    #
    def write_series # :nodoc:
      write_series_base { nil }
    end

    def write_series_base
      # Write each series with subelements.
      index = 0
      @series.each do |series|
        write_ser(index, series)
        index += 1
      end

      # Write the c:marker element.
      write_marker_value

      # Write the c:overlap element
      # block given by Bar and Column
      yield

      # Generate the axis ids.
      add_axis_id
      add_axis_id

      # Write the c:axId element.
      write_axis_id(@axis_ids[0])
      write_axis_id(@axis_ids[1])
    end

    #
    # Write the <c:ser> element.
    #
    def write_ser(index, series) # :nodoc:
      @writer.tag_elements('c:ser') do
        # Write the c:idx element.
        write_idx(index)
        # Write the c:order element.
        write_order(index)
        # Write the series name.
        write_series_name(series)
        # Write the c:spPr element.
        write_sp_pr(series)
        # Write the c:marker element.
        write_marker(series[:_marker])
        # Write the c:invertIfNegative element.
        write_c_invert_if_negative(series[:_invert_if_neg])
        # Write the c:dLbls element.
        write_d_lbls(series[:labels])
        # Write the c:trendline element.
        write_trendline(series[:trendline])
        # Write the c:cat element.
        write_cat(series)
        # Write the c:val element.
        write_val(series)
      end
    end

    #
    # Write the <c:idx> element.
    #
    def write_idx(val) # :nodoc:
      attributes = ['val', val]

      @writer.empty_tag('c:idx', attributes)
    end

    #
    # Write the <c:order> element.
    #
    def write_order(val) # :nodoc:
      attributes = ['val', val]

      @writer.empty_tag('c:order', attributes)
    end

    #
    # Write the series name.
    #
    def write_series_name(series) # :nodoc:
      if name = series[:_name_formula]
        write_tx_formula(name, series[:_name_id])
      elsif name = series[:_name]
        write_tx_value(name)
      end
    end

    #
    # Write the <c:cat> element.
    #
    def write_cat(series) # :nodoc:

      formula = series[:_categories]
      data_id = series[:_cat_data_id]

      data = @formula_data[data_id] if data_id

      # Ignore <c:cat> elements for charts without category values.
      return unless formula

      @has_category = 1

      @writer.tag_elements('c:cat') do
        # Check the type of cached data.
        type = get_data_type(data)
        if type == 'str'
          @has_category = 0
          # Write the c:numRef element.
          write_str_ref(formula, data, type)
        else
          # Write the c:numRef element.
          write_num_ref(formula, data, type)
        end
      end
    end

    #
    # Write the <c:val> element.
    #
    def write_val(series) # :nodoc:
      write_val_base(series[:_values], series[:_val_data_id], 'c:val')
    end

    def write_val_base(formula, data_id, tag) # :nodoc:
      data    = @formula_data[data_id]

      @writer.tag_elements(tag) do
        # Check the type of cached data.
        type = get_data_type(data)
        if type == 'str'
          # Write the c:numRef element.
          write_str_ref(formula, data, type)
        else
          # Write the c:numRef element.
          write_num_ref(formula, data, type)
        end
      end
    end

    #
    # Write the <c:numRef> element.
    #
    def write_num_ref(formula, data, type) # :nodoc:
      @writer.tag_elements('c:numRef') do
        # Write the c:f element.
        write_series_formula(formula)
        if type == 'num'
          # Write the c:numCache element.
          write_num_cache(data)
        elsif type == 'str'
          # Write the c:strCache element.
          write_str_cache(data)
        end
      end
    end

    #
    # Write the <c:strRef> element.
    #
    def write_str_ref(formula, data, type) # :nodoc:
      @writer.tag_elements('c:strRef') do
        # Write the c:f element.
        write_series_formula(formula)
        if type == 'num'
          # Write the c:numCache element.
          write_num_cache(data)
        elsif type == 'str'
          # Write the c:strCache element.
          write_str_cache(data)
        end
      end
    end

    #
    # Write the <c:f> element.
    #
    def write_series_formula(formula) # :nodoc:
      # Strip the leading '=' from the formula.
      formula = formula.sub(/^=/, '')

      @writer.data_element('c:f', formula)
    end

    #
    # Write the <c:axId> element.
    #
    def write_axis_id(val) # :nodoc:
      attributes = ['val', val]

      @writer.empty_tag('c:axId', attributes)
    end

    #
    # Write the <c:catAx> element. Usually the X axis.
    #
    def write_cat_axis(position = nil) # :nodoc:
      position = @cat_axis_position
      horiz    = @horiz_cat_axis
      x_axis   = @x_axis
      y_axis   = @y_axis

      # Overwrite the default axis position with a user supplied value.
      position = x_axis[:_position] || position

      @writer.tag_elements('c:catAx') do
        write_axis_id(@axis_ids[0])
        # Write the c:scaling element.
        write_scaling(x_axis[:_reverse])
        # Write the c:axPos element.
        write_axis_pos(position, y_axis[:_reverse])
        # Write the axis title elements.
        if title = @x_axis[:_formula]
          write_title_formula(title, @x_axis[:_data_id], horiz)
        elsif title = @x_axis[:_name]
          write_title_rich(title, horiz)
        end
        # Write the c:numFmt element.
        write_num_fmt
        # Write the c:tickLblPos element.
        write_tick_label_pos(x_axis[:_label_position])
        # Write the c:crossAx element.
        write_cross_axis(@axis_ids[1])
        # Note, the category crossing comes from the value axis.
        if nil_or_max?(y_axis[:_crossing])
          # Write the c:crosses element.
          write_crosses(y_axis[:_crossing])
        else
          # Write the c:crossesAt element.
          write_c_crosses(y_axis[:_crossing])
        end
        # Write the c:auto element.
        write_auto(1)
        # Write the c:labelAlign element.
        write_label_align('ctr')
        # Write the c:labelOffset element.
        write_label_offset(100)
      end
    end

    #
    # Write the <c:valAx> element. Usually the Y axis.
    #
    # TODO. Maybe should have a _write_cat_val_axis method as well for scatter.
    #
    def write_val_axis(position = nil, hide_major_gridlines = nil) # :nodoc:
      params = {
        :axis_position  => @y_axis[:_position],
        :axis_id        => @axis_ids[1],
        :scaling_axis   => @y_axis,
        :axis_position_element => @x_axis[:_reverse],
        :title_axis     => @y_axis,
        :tick_label_pos => @y_axis[:_label_position],
        :cross_axis     => @axis_ids[0],
        :category_crossing => @x_axis[:_crossing],
        :major_unit     => @y_axis[:_major_unit],
        :minor_unit     => @y_axis[:_minor_unit]
      }
      write_val_axis_common(position, hide_major_gridlines, params)
    end

    #
    # Write the <c:valAx> element. This is for the second valAx in scatter plots.
    #
    # Usually the X axis.
    #
    def write_cat_val_axis(position, hide_major_gridlines) # :nodoc:
      params = {
        :axis_position  => @x_axis[:_position],
        :axis_id        => @axis_ids[0],
        :scaling_axis   => @x_axis,
        :axis_position_element => @y_axis[:_reverse],
        :title_axis     => @x_axis,
        :tick_label_pos => @x_axis[:_label_position],
        :cross_axis     => @axis_ids[1],
        :category_crossing => @y_axis[:_crossing],
        :major_unit     => @x_axis[:_major_unit],
        :minor_unit     => @x_axis[:_minor_unit]
      }
      write_val_axis_common(position, hide_major_gridlines, params)
    end

    def write_val_axis_common(position, hide_major_gridlines, params) # :nodoc:
      position ||= @val_axis_position
      horiz      = @horiz_val_axis

      # Overwrite the default axis position with a user supplied value.
      position = params[:axis_position] || position

      @writer.tag_elements('c:valAx') do
        write_axis_id(params[:axis_id])
        # Write the c:scaling element.
        write_scaling(
                      params[:scaling_axis][:_reverse], params[:scaling_axis][:_min],
                      params[:scaling_axis][:_max], params[:scaling_axis][:_log_base])
        # Write the c:axPos element.
        write_axis_pos(position, params[:axis_position_element])
        # Write the c:majorGridlines element.
        write_major_gridlines unless hide_major_gridlines
        # Write the axis title elements.
        if title = params[:title_axis][:_formula]
          write_title_formula(title, @y_axis[:_data_id], horiz)
        elsif title = params[:title_axis][:_name]
          write_title_rich(title, horiz)
        end
        # Write the c:numberFormat element.
        write_number_format
        # Write the c:tickLblPos element.
        write_tick_label_pos(params[:tick_label_pos])
        # Write the c:crossAx element.
        write_cross_axis(params[:cross_axis])
        # Note, the category crossing comes from the value axis.
        if nil_or_max?(params[:category_crossing])
          # Write the c:crosses element.
          write_crosses(params[:category_crossing])
        else
          # Write the c:crossesAt element.
          write_c_crosses_at(params[:category_crossing])
        end
        # Write the c:crossBetween element.
        write_cross_between
        # Write the c:majorUnit element.
        write_c_major_unit(params[:major_unit])
        # Write the c:minorUnit element.
        write_c_minor_unit(params[:minor_unit])
      end
    end

    #
    # Write the <c:dateAx> element. Usually the X axis.
    #
    def write_date_axis(position = nil) # :nodoc:
      position  = @cat_axis_position
      x_axis    = @x_axis
      y_axis    = @y_axis

      @writer.tag_elements('c:dateAx') do
        write_axis_id(@axis_ids[0])
        # Write the c:scaling element.
        write_scaling(x_axis[:reverse], x_axis[:_min], x_axis[:_max], x_axis[:_log_base])
        # Write the c:axPos element.
        write_axis_pos(position, y_axis[:reverse])
        # Write the axis title elements.
        if title = x_axis[:_formula]
          write_title_formula(title, x_axis[:_data_id])
        elsif title = x_axis[:_name]
          write_title_rich(title)
        end
        # Write the c:numFmt element.
        write_num_fmt('dd/mm/yyyy')
        # Write the c:tickLblPos element.
        write_tick_label_pos(x_axis[:_label_position])
        # Write the c:crossAx element.
        write_cross_axis(@axis_ids[1])
        # Note, the category crossing comes from the value axis.
        if nil_or_max?(y_axis[:_crossing])
          # Write the c:crossing element.
          write_crosses(y_axis[:_crossing])
        else
          # Write the c:crossesAt element.
          write_c_crosses_at(y_axis[:_crossing])
        end
        # Write the c:auto element.
        write_auto(1)
        # Write the c:labelOffset element.
        write_label_offset(100)
        # Write the c:majorUnit element.
        write_c_major_unit(x_axis[:_major_unit])
        # Write the c:majorTimeUnit element.
        if !x_axis[:_major_unit].nil?
          write_c_major_time_unit(x_axis[:_major_unit_type])
        end
        # Write the c:minorUnit element.
        write_c_minor_unit(x_axis[:_minor_unit])
        # Write the c:minorTimeUnit element.
        if !x_axis[:_minor_unit].nil?
          write_c_minor_time_unit(x_axis[:_minor_unit_type])
        end
      end
    end

    #
    # Write the <c:scaling> element.
    #
    def write_scaling(reverse, min = nil, max = nil, log_base = nil) # :nodoc:
      @writer.tag_elements('c:scaling') do
        # Write the c:logBase element.
        write_c_log_base(log_base)
        # Write the c:orientation element.
        write_orientation(reverse)
        # Write the c:max element.
        write_c_max(max)
        # Write the c:min element.
        write_c_min(min)
      end
    end

    #
    # Write the <c:orientation> element.
    #
    def write_orientation(reverse = nil) # :nodoc:
      val     = reverse ? 'maxMin' : 'minMax'

      attributes = ['val', val]

      @writer.empty_tag('c:orientation', attributes)
    end

    #
    # Write the <c:logBase> element.
    #
    def write_c_log_base(val) # :nodoc:
      return if val == 0 || val.nil?

      attributes = ['val', val]

      @writer.empty_tag('c:logBase', attributes)
    end

    #
    # Write the <c:orientation> element.
    #
    def write_orientation(reverse = 'maxMin') # :nodoc:
      val = 'minMax'

      attributes = ['val', val]

      @writer.empty_tag('c:orientation', attributes)
    end

    #
    # Write the <c:max> element.
    #
    def write_c_max(max = nil) # :nodoc:
      return if max.nil?

      attributes = ['val', max]

      @writer.empty_tag('c:max', attributes)
    end

    #
    # Write the <c:min> element.
    #
    def write_c_min(min = nil) # :nodoc:
      return if min.nil?

      attributes = ['val', min]

      @writer.empty_tag('c:min', attributes)
    end

    #
    # Write the <c:axPos> element.
    #
    def write_axis_pos(val, reverse = false) # :nodoc:
      if reverse
        val = 'r' if val == 'l'
        val = 't' if val == 'b'
      end

      attributes = ['val', val]

      @writer.empty_tag('c:axPos', attributes)
    end

    #
    # Write the <c:numFmt> element.
    #
    def write_num_fmt(format_code = nil) # :nodoc:
      format_code ||= 'General'
      source_linked = 1

      # These elements are only required for charts with categories.
      return unless @has_category

      attributes = [
                    'formatCode',   format_code,
                    'sourceLinked', source_linked
                   ]

      @writer.empty_tag('c:numFmt', attributes)
    end

    #
    # Write the <c:tickLblPos> element.
    #
    def write_tick_label_pos(val) # :nodoc:
      val ||= 'nextTo'
      val = 'nextTo' if val == 'next_to'

      attributes = ['val', val]

      @writer.empty_tag('c:tickLblPos', attributes)
    end

    #
    # Write the <c:crossAx> element.
    #
    def write_cross_axis(val = 'autoZero') # :nodoc:
      attributes = ['val', val]

      @writer.empty_tag('c:crossAx', attributes)
    end

    #
    # Write the <c:crosses> element.
    #
    def write_crosses(val) # :nodoc:
      val ||= 'autoZero'

      attributes = ['val', val]

      @writer.empty_tag('c:crosses', attributes)
    end

    #
    # Write the <c:crossesAt> element.
    #
    def write_c_crosses_at(val) # :nodoc:
      attributes = ['val', val]

      @writer.empty_tag('c:crossesAt', attributes)
    end

    #
    # Write the <c:auto> element.
    #
    def write_auto(val) # :nodoc:
      attributes = ['val', val]

      @writer.empty_tag('c:auto', attributes)
    end

    #
    # Write the <c:labelAlign> element.
    #
    def write_label_align(val) # :nodoc:
      attributes = ['val', val]

      @writer.empty_tag('c:lblAlgn', attributes)
    end

    #
    # Write the <c:labelOffset> element.
    #
    def write_label_offset(val) # :nodoc:
      attributes = ['val', val]

      @writer.empty_tag('c:lblOffset', attributes)
    end

    #
    # Write the <c:majorGridlines> element.
    #
    def write_major_gridlines # :nodoc:
      @writer.empty_tag('c:majorGridlines')
    end

    #
    # Write the <c:numberFormat> element.
    #
    # TODO. Merge/replace with _write_num_fmt.
    #
    def write_number_format # :nodoc:
      format_code   = 'General'
      source_linked = 1

      attributes = [
                    'formatCode',   format_code,
                    'sourceLinked', source_linked
                   ]

      @writer.empty_tag('c:numFmt', attributes)
    end

    #
    # Write the <c:crossBetween> element.
    #
    def write_cross_between # :nodoc:
      val  = @cross_between || 'between'

      attributes = ['val', val]

      @writer.empty_tag('c:crossBetween', attributes)
    end

    #
    # Write the <c:majorUnit> element.
    #
    def write_c_major_unit(val = nil) # :nodoc:
      return unless val

      attributes = ['val', val]

      @writer.empty_tag('c:majorUnit', attributes)
    end

    #
    # Write the <c:minorUnit> element.
    #
    def write_c_minor_unit(val = nil) # :nodoc:
      return unless val

      attributes = ['val', val]

      @writer.empty_tag('c:minorUnit', attributes)
    end

    #
    # Write the <c:majorTimeUnit> element.
    #
    def write_c_major_time_unit(val = 'days') # :nodoc:
      attributes = ['val', val]

      @writer.empty_tag('c:majorTimeUnit', attributes)
    end

    #
    # Write the <c:minorTimeUnit> element.
    #
    def write_c_minor_time_unit(val = 'days') # :nodoc:
      attributes = ['val', val]

      @writer.empty_tag('c:minorTimeUnit', attributes)
    end

    #
    # Write the <c:legend> element.
    #
    def write_legend # :nodoc:
      position      = @legend_position
      overlay       = false

      if @legend_delete_series && @legend_delete_series.kind_of?(Array)
        @delete_series = @legend_delete_series
      end

      if position =~ /^overlay_/
        position.sub!(/^overlay_/, '')
        overlay = true if position
      end

      allowed = {
        'right'  => 'r',
        'left'   => 'l',
        'top'    => 't',
        'bottom' => 'b'
      }

      return if position == 'none'
      return unless allowed.has_key?(position)

      position = allowed[position]

      @writer.tag_elements('c:legend') do
        # Write the c:legendPos element.
        write_legend_pos(position)
        # Remove series labels from the legend.
        @delete_series.each do |index|
          # Write the c:legendEntry element.
          write_legend_entry(index)
        end if @delete_series
        # Write the c:layout element.
        write_layout
        # Write the c:overlay element.
        write_overlay if overlay
      end
    end

    #
    # Write the <c:legendPos> element.
    #
    def write_legend_pos(val) # :nodoc:
      attributes = ['val', val]

      @writer.empty_tag('c:legendPos', attributes)
    end

    #
    # Write the <c:legendEntry> element.
    #
    def write_legend_entry(index) # :nodoc:
      @writer.tag_elements('c:legendEntry') do
        # Write the c:idx element.
        write_idx(index)
        # Write the c:delete element.
        write_delete(1)
      end
    end

    #
    # Write the <c:overlay> element.
    #
    def write_overlay # :nodoc:
      val  = 1

      attributes = ['val', val]

      @writer.empty_tag('c:overlay', attributes)
    end

    #
    # Write the <c:plotVisOnly> element.
    #
    def write_plot_vis_only # :nodoc:
      val  = 1

      attributes = ['val', val]

      @writer.empty_tag('c:plotVisOnly', attributes)
    end

    #
    # Write the <c:printSettings> element.
    #
    def write_print_settings # :nodoc:
      @writer.tag_elements('c:printSettings') do
        # Write the c:headerFooter element.
        write_header_footer
        # Write the c:pageMargins element.
        write_page_margins
        # Write the c:pageSetup element.
        write_page_setup
      end
    end

    #
    # Write the <c:headerFooter> element.
    #
    def write_header_footer # :nodoc:
      @writer.empty_tag('c:headerFooter')
    end

    #
    # Write the <c:pageMargins> element.
    #
    def write_page_margins # :nodoc:
      b      = 0.75
      l      = 0.7
      r      = 0.7
      t      = 0.75
      header = 0.3
      footer = 0.3

      attributes = [
                    'b',      b,
                    'l',      l,
                    'r',      r,
                    't',      t,
                    'header', header,
                    'footer', footer
                   ]

      @writer.empty_tag('c:pageMargins', attributes)
    end

    #
    # Write the <c:pageSetup> element.
    #
    def write_page_setup # :nodoc:
      @writer.empty_tag('c:pageSetup')
    end

    #
    # Write the <c:title> element for a rich string.
    #
    def write_title_rich(title, horiz = nil) # :nodoc:
      @writer.tag_elements('c:title') do
        # Write the c:tx element.
        write_tx_rich(title, horiz)
        # Write the c:layout element.
        write_layout
      end
    end

    #
    # Write the <c:title> element for a rich string.
    #
    def write_title_formula(title, data_id, horiz) # :nodoc:
      @writer.tag_elements('c:title') do
        # Write the c:tx element.
        write_tx_formula(title, data_id)
        # Write the c:layout element.
        write_layout
        # Write the c:txPr element.
        write_tx_pr(horiz)
      end
    end

    #
    # Write the <c:tx> element.
    #
    def write_tx_rich(title, horiz) # :nodoc:
      @writer.tag_elements('c:tx') { write_rich(title, horiz) }
    end

    #
    # Write the <c:tx> element with a simple value such as for series names.
    #
    def write_tx_value(title) # :nodoc:
      @writer.tag_elements('c:tx') { write_v(title) }
    end

    #
    # Write the <c:tx> element.
    #
    def write_tx_formula(title, data_id) # :nodoc:
      data = @formula_data[data_id] if data_id

      @writer.tag_elements('c:tx') { write_str_ref(title, data, 'str') }
    end

    #
    # Write the <c:rich> element.
    #
    def write_rich(title, horiz) # :nodoc:
      @writer.tag_elements('c:rich') do
        # Write the a:bodyPr element.
        write_a_body_pr(horiz)
        # Write the a:lstStyle element.
        write_a_lst_style
        # Write the a:p element.
        write_a_p_rich(title)
      end
    end

    #
    # Write the <a:bodyPr> element.
    #
    def write_a_body_pr(horiz) # :nodoc:
      rot   = -5400000
      vert  = 'horz'

      attributes = [
                    'rot',  rot,
                    'vert', vert
                   ]

      attributes = [] if !horiz || horiz == 0

      @writer.empty_tag('a:bodyPr', attributes)
    end

    #
    # Write the <a:lstStyle> element.
    #
    def write_a_lst_style # :nodoc:
      @writer.empty_tag('a:lstStyle')
    end

    #
    # Write the <a:p> element for rich string titles.
    #
    def write_a_p_rich(title) # :nodoc:
      @writer.tag_elements('a:p') do
        # Write the a:pPr element.
        write_a_p_pr_rich
        # Write the a:r element.
        write_a_r(title)
      end
    end

    #
    # Write the <a:p> element for formula titles.
    #
    def write_a_p_formula(title) # :nodoc:
      @writer.tag_elements('a:p') do
        # Write the a:pPr element.
        write_a_p_pr_formula
        # Write the a:endParaRPr element.
        write_a_end_para_rpr
      end
    end

    #
    # Write the <a:pPr> element for rich string titles.
    #
    def write_a_p_pr_rich # :nodoc:
      @writer.tag_elements('a:pPr') { write_a_def_rpr }
    end

    #
    # Write the <a:pPr> element for formula titles.
    #
    def write_a_p_pr_formula # :nodoc:
      @writer.tag_elements('a:pPr') { write_a_def_rpr }
    end

    #
    # Write the <a:defRPr> element.
    #
    def write_a_def_rpr # :nodoc:
      @writer.empty_tag('a:defRPr')
    end

    #
    # Write the <a:endParaRPr> element.
    #
    def write_a_end_para_rpr # :nodoc:
      lang = 'en-US'

      attributes = ['lang', lang]

      @writer.empty_tag('a:endParaRPr', attributes)
    end

    #
    # Write the <a:r> element.
    #
    def write_a_r(title) # :nodoc:
      @writer.tag_elements('a:r') do
        # Write the a:rPr element.
        write_a_r_pr
        # Write the a:t element.
        write_a_t(title)
      end
    end

    #
    # Write the <a:rPr> element.
    #
    def write_a_r_pr # :nodoc:
      lang = 'en-US'

      attributes = ['lang', lang]

      @writer.empty_tag('a:rPr', attributes)
    end

    #
    # Write the <a:t> element.
    #
    def write_a_t(title) # :nodoc:
      @writer.data_element('a:t', title)
    end

    #
    # Write the <c:txPr> element.
    #
    def write_tx_pr(horiz) # :nodoc:
      @writer.tag_elements('c:txPr') do
        # Write the a:bodyPr element.
        write_a_body_pr(horiz)
        # Write the a:lstStyle element.
        write_a_lst_style
        # Write the a:p element.
        write_a_p_formula
      end
    end

    #
    # Write the <c:marker> element.
    #
    def write_marker(marker = nil) # :nodoc:
      marker ||= @default_marker

      return if marker.nil? || marker == 0
      return if marker[:automatic] && marker[:automatic] != 0

      @writer.tag_elements('c:marker') do
        # Write the c:symbol element.
        write_symbol(marker[:type])
        # Write the c:size element.
        size = marker[:size]
        write_marker_size(size) if !size.nil? && size != 0
        # Write the c:spPr element.
        write_sp_pr(marker)
      end
    end

    #
    # Write the <c:marker> element without a sub-element.
    #
    def write_marker_value # :nodoc:
      style = @default_marker

      return unless style

      attributes = ['val', 1]

      @writer.empty_tag('c:marker', attributes)
    end

    #
    # Write the <c:size> element.
    #
    def write_marker_size(val) # :nodoc:
      attributes = ['val', val]

      @writer.empty_tag('c:size', attributes)
    end

    #
    # Write the <c:symbol> element.
    #
    def write_symbol(val) # :nodoc:
      attributes = ['val', val]

      @writer.empty_tag('c:symbol', attributes)
    end

    #
    # Write the <c:spPr> element.
    #
    def write_sp_pr(series) # :nodoc:
      return if (!series.has_key?(:_line) || series[:_line][:_defined].nil? || series[:_line][:_defined]== 0) &&
                (!series.has_key?(:_fill) || series[:_fill][:_defined].nil? || series[:_fill][:_defined]== 0)

      @writer.tag_elements('c:spPr') do
        # Write the a:solidFill element for solid charts such as pie and bar.
        write_a_solid_fill(series[:_fill]) if series[:_fill] && series[:_fill][:_defined] != 0
        # Write the a:ln element.
        write_a_ln(series[:_line]) if series[:_line] && series[:_line][:_defined]
      end
    end

    #
    # Write the <a:ln> element.
    #
    def write_a_ln(line) # :nodoc:
      attributes = []

      # Add the line width as an attribute.
      if width = line[:width]
        # Round width to nearest 0.25, like Excel.
        width = ((width + 0.125) * 4).to_i / 4.0

        # Convert to internal units.
        width = (0.5 + (12700 * width)).to_i

        attributes = ['w', width]
      end

      @writer.tag_elements('a:ln', attributes) do
        # Write the line fill.
        if !line[:none].nil? && line[:none] != 0
          # Write the a:noFill element.
          write_a_no_fill
        else
          # Write the a:solidFill element.
          write_a_solid_fill(line)
        end
        # Write the line/dash type.
        if type = line[:dash_type]
          # Write the a:prstDash element.
          write_a_prst_dash(type)
        end
      end
    end

    #
    # Write the <a:noFill> element.
    #
    def write_a_no_fill # :nodoc:
      @writer.empty_tag('a:noFill')
    end

    #
    # Write the <a:solidFill> element.
    #
    def write_a_solid_fill(line) # :nodoc:
      @writer.tag_elements('a:solidFill') do
        if line[:color]
          color = get_color(line[:color])

          # Write the a:srgbClr element.
          write_a_srgb_clr(color)
        end
      end
    end

    #
    # Write the <a:srgbClr> element.
    #
    def write_a_srgb_clr(val) # :nodoc:
      attributes = ['val', val]

      @writer.empty_tag('a:srgbClr', attributes)
    end

    #
    # Write the <a:prstDash> element.
    #
    def write_a_prst_dash(val) # :nodoc:
      attributes = ['val', val]

      @writer.empty_tag('a:prstDash', attributes)
    end

    #
    # Write the <c:trendline> element.
    #
    def write_trendline(trendline) # :nodoc:
      return unless trendline

      @writer.tag_elements('c:trendline') do
        # Write the c:name element.
        write_name(trendline[name])
        # Write the c:spPr element.
        write_sp_pr(trendline)
        # Write the c:trendlineType element.
        write_trendline_type(trendline[type])
        # Write the c:order element for polynomial trendlines.
        write_trendline_order(trendline[order]) if trendline[type] == 'poly'
        # Write the c:period element for moving average trendlines.
        write_period(trendline[period]) if trendline[type] == 'movingAvg'
        # Write the c:forward element.
        write_forward(trendline[forward])
        # Write the c:backward element.
        write_backward(trendline[backward])
      end
    end

    #
    # Write the <c:trendlineType> element.
    #
    def write_trendline_type(val) # :nodoc:
      attributes = ['val', val]

      @writer.empty_tag('c:trendlineType', attributes)
    end

    #
    # Write the <c:name> element.
    #
    def write_name(data) # :nodoc:
      return unless data

      @writer.data_element('c:name', data)
    end

    #
    # Write the <c:order> element.
    #
    def write_trendline_order(val = 2) # :nodoc:
      attributes = ['val', val]

      @writer.empty_tag('c:order', attributes)
    end

    #
    # Write the <c:period> element.
    #
    def write_period(val = 2) # :nodoc:
      attributes = ['val', val]

      @writer.empty_tag('c:period', attributes)
    end

    #
    # Write the <c:forward> element.
    #
    def write_forward(val) # :nodoc:
      return unless val

      attributes = ['val', val]

      @writer.empty_tag('c:forward', attributes)
    end

    #
    # Write the <c:backward> element.
    #
    def write_backward(val) # :nodoc:
      return unless val

      attributes = ['val', val]

      @writer.empty_tag('c:backward', attributes)
    end

    #
    # Write the <c:hiLowLines> element.
    #
    def write_hi_low_lines # :nodoc:
      @writer.empty_tag('c:hiLowLines')
    end

    #
    # Write the <c:overlap> element.
    #
    def write_overlap # :nodoc:
      val  = 100

      attributes = ['val', val]

      @writer.empty_tag('c:overlap', attributes)
    end

    #
    # Write the <c:numCache> element.
    #
    def write_num_cache(data) # :nodoc:
      @writer.tag_elements('c:numCache') do
        write_format_code('General')
        write_pt_count(data.size)
        write_pts(data)
      end
    end

    #
    # Write the <c:strCache> element.
    #
    def write_str_cache(data) # :nodoc:
      @writer.tag_elements('c:strCache') do
        write_pt_count(data.size)
        write_pts(data)
      end
    end

    def write_pts(data)
      data.each_index { |i| write_pt(i, data[i])}
    end

    #
    # Write the <c:formatCode> element.
    #
    def write_format_code(data) # :nodoc:
      @writer.data_element('c:formatCode', data)
    end

    #
    # Write the <c:ptCount> element.
    #
    def write_pt_count(val) # :nodoc:
      attributes = ['val', val]

      @writer.empty_tag('c:ptCount', attributes)
    end

    #
    # Write the <c:pt> element.
    #
    def write_pt(idx, value) # :nodoc:
      return unless value

      attributes = ['idx', idx]

      @writer.tag_elements('c:pt', attributes) { write_v(value) }
    end

    #
    # Write the <c:v> element.
    #
    def write_v(data) # :nodoc:
      @writer.data_element('c:v', data)
    end

    #
    # Write the <c:protection> element.
    #
    def write_protection # :nodoc:
      return if @protection == 0

      @writer.empty_tag('c:protection')
    end

    #
    # Write the <c:dLbls> element.
    #
    def write_d_lbls(labels) # :nodoc:
      return unless labels

      @writer.tag_elements('c:dLbls') do
        # Write the c:showVal element.
        write_show_val if labels[value]
        # Write the c:showCatName element.
        write_show_cat_name if labels[category]
        # Write the c:showSerName element.
        write_show_ser_name if labels[series_name]
      end
    end

    #
    # Write the <c:showVal> element.
    #
    def write_show_val # :nodoc:
      val  = 1

      attributes = ['val', val]

      @writer.empty_tag('c:showVal', attributes)
    end

    #
    # Write the <c:showCatName> element.
    #
    def write_show_cat_name # :nodoc:
      val  = 1

      attributes = ['val', val]

      @writer.empty_tag('c:showCatName', attributes)
    end

    #
    # Write the <c:showSerName> element.
    #
    def write_show_ser_name # :nodoc:
      val  = 1

      attributes = ['val', val]

      @writer.empty_tag('c:showSerName', attributes)
    end

    #
    # Write the <c:delete> element.
    #
    def write_delete(val) # :nodoc:
      attributes = ['val', val]

      @writer.empty_tag('c:delete', attributes)
    end

    #
    # Write the <c:invertIfNegative> element.
    #
    def write_c_invert_if_negative(invert = nil) # :nodoc:
      val    = 1

      return unless invert

      attributes = ['val', val]

      @writer.empty_tag('c:invertIfNegative', attributes)
    end

    def nil_or_max?(val)
      val.nil? || val == 'max'
    end
  end
end
