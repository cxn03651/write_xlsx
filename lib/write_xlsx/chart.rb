# -*- coding: utf-8 -*-
require 'write_xlsx/package/xml_writer_simple'
require 'write_xlsx/utility'
require 'write_xlsx/chart/axis'

module Writexlsx
  class Table
    include Writexlsx::Utility

    attr_reader :horizontal, :vertical, :outline, :show_keys

    def initialize(params = {})
      @horizontal, @vertical, @outline, @show_keys = true, true, true, false
      @horizontal = params[:horizontal] if params.has_key?(:horizontal)
      @vertical   = params[:vertical]   if params.has_key?(:vertical)
      @outline    = params[:outline]    if params.has_key?(:outline)
      @show_keys  = params[:show_keys]  if params.has_key?(:show_keys)
    end

    def write_d_table(writer)
      writer.tag_elements('c:dTable') do
        writer.empty_tag('c:showHorzBorder', attributes) if ptrue?(horizontal)
        writer.empty_tag('c:showVertBorder', attributes) if ptrue?(vertical)
        writer.empty_tag('c:showOutline',    attributes) if ptrue?(outline)
        writer.empty_tag('c:showKeys',       attributes) if ptrue?(show_keys)
      end
    end

    private

    def attributes
      ['val', 1]
    end
  end

  # ==SYNOPSIS
  #
  # To create a simple Excel file with a chart using WriteXLSX:
  #
  #     require 'rubygems'
  #     require 'write_xlsx'
  #
  #     workbook  = WriteXLSX.new('chart.xlsx')
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
  # The Chart is an abstract base class for modules that implement
  # charts in WriteXLSX. The information below is applicable to all of
  # the available subclasses.
  #
  # The Chart isn't used directly. A chart object is created via
  # the {WriteXLXS#add_chart()}
  # method where the chart type is specified:
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
  # Creates a Column style (histogram) chart. See Writexlsx::Chart::Column.
  #
  # ===line
  # Creates a Line style chart. See Writexlsx::Chart::Line.
  #
  # ===pie
  # Creates a Pie style chart. See Writexlsx::Chart::Pie.
  #
  # ===scatter
  # Creates a Scatter style chart. See Writexlsx::Chart::Scatter.
  #
  # ===stock
  # Creates a Stock style chart. See Writexlsx::Chart::Stock.
  #
  # ===radar
  # Creates a Radar style chart. See Writexlsx::Chart::Radar.
  #
  # Chart subtypes are also supported in some cases:
  #
  #     workbook.add_chart(:type => 'bar', :subtype => 'stacked')
  #
  # The currently available subtypes are:
  #
  #     area
  #         stacked
  #         percent_stacked
  #
  #     bar
  #         stacked
  #         percent_stacked
  #
  #     column
  #         stacked
  #         percent_stacked
  #
  #     scatter
  #         straight_with_markers
  #         straight
  #         smooth_with_markers
  #         smooth
  #
  #     radar
  #         with_markers
  #         filled
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
  #
  # Chart formatting properties are generally set using hash refs.
  #
  #     chart.add_series(
  #         :values     => '=Sheet1!$B$1:$B$5',
  #         :line       => { color => 'blue' }
  #     )
  #
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
  #
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
  #     :value
  #     :category
  #     :series_name
  #     :position
  #     :leader_lines
  #     :percentage
  #
  # The value property turns on the Value data label for a series.
  #
  #     chart.add_series(
  #         :values      => '=Sheet1!$B$1:$B$5',
  #         :data_labels => { :value => 1 }
  #     )
  # The category property turns on the Category Name data label for a series.
  #
  #     chart.add_series(
  #         :values      => '=Sheet1!$B$1:$B$5',
  #         :data_labels => { :category => 1 }
  #     )
  # The series_name property turns on the Series Name data label for a series.
  #
  #     chart.add_series(
  #         :values      => '=Sheet1!$B$1:$B$5',
  #         :data_labels => { :series_name => 1 }
  #     )
  # The C<position> property is used to position the data label for a series.
  #
  #     chart.add_series(
  #         :values      => '=Sheet1!$B$1:$B$5',
  #         :data_labels => { :value => 1, :position => 'center' }
  #     )
  #
  # Valid positions are:
  #
  #     :center
  #     :right
  #     :left
  #     :top
  #     :bottom
  #     :above           # Same as top
  #     :below           # Same as bottom
  #     :inside_end      # Pie chart mainly.
  #     :outside_end     # Pie chart mainly.
  #     :best_fit        # Pie chart mainly.
  #
  # The C<percentage> property is used to turn on the I<Percentage>
  # for the data label for a series. It is mainly used for pie charts.
  #
  #    chart.add_series(
  #         :values      => '=Sheet1!$B$1:$B$5',
  #         :data_labels => { :percentage => 1 }
  #    )
  #
  # The C<leader_lines> property is used to turn on  I<Leader Lines>
  # for the data label for a series. It is mainly used for pie charts.
  #
  #    chart.add_series(
  #         :values      => '=Sheet1!$B$1:$B$5',
  #         :data_labels => { :value => 1, :leader_lines => 1 }
  #    )
  #
  class Chart
    include Writexlsx::Utility

    attr_accessor :id, :name   # :nodoc:
    attr_writer :index, :palette, :protection   # :nodoc:
    attr_reader :embedded, :formula_ids, :formula_data   # :nodoc:
    attr_reader :x_scale, :y_scale, :x_offset, :y_offset # :nodoc:
    attr_reader :width, :height  # :nodoc:

    #
    # Factory method for returning chart objects based on their class type.
    #
    def self.factory(current_subclass, subtype = nil) # :nodoc:
      case current_subclass.downcase.capitalize
      when 'Area'
        require 'write_xlsx/chart/area'
        Chart::Area.new(subtype)
      when 'Bar'
        require 'write_xlsx/chart/bar'
        Chart::Bar.new(subtype)
      when 'Column'
        require 'write_xlsx/chart/column'
        Chart::Column.new(subtype)
      when 'Line'
        require 'write_xlsx/chart/line'
        Chart::Line.new(subtype)
      when 'Pie'
        require 'write_xlsx/chart/pie'
        Chart::Pie.new(subtype)
      when 'Radar'
        require 'write_xlsx/chart/radar'
        Chart::Radar.new(subtype)
      when 'Scatter'
        require 'write_xlsx/chart/scatter'
        Chart::Scatter.new(subtype)
      when 'Stock'
        require 'write_xlsx/chart/stock'
        Chart::Stock.new(subtype)
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
      @series_index      = 0
      @style_id          = 2
      @axis_ids          = []
      @axis2_ids         = []
      @cat_has_num_fmt   = false
      @requires_category = 0
      @legend_position   = 'right'
      @cat_axis_position = 'b'
      @val_axis_position = 'l'
      @formula_ids       = {}
      @formula_data      = []
      @horiz_cat_axis    = 0
      @horiz_val_axis    = 1
      @protection        = 0
      @chartarea         = {}
      @plotarea          = {}
      @x_axis            = Axis.new
      @y_axis            = Axis.new
      @x2_axis           = Axis.new
      @y2_axis           = Axis.new
      @name              = ''
      @show_blanks       = 'gap'
      @show_hidden_data  = false
      @show_crosses      = true
      @width             = 480
      @height            = 288
      @x_scale           = 1
      @y_scale           = 1
      @x_offset          = 0
      @y_offset          = 0
      @table             = nil
      @smooth_allowed    = 0

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
      write_chart_space do

        # Write the c:lang element.
        write_lang

        # Write the c:style element.
        write_style

        # Write the c:protection element.
        write_protection

        # Write the c:chart element.
        write_chart

        # Write the c:spPr element for the chartarea formatting.
        write_sp_pr(@chartarea)

        # Write the c:printSettings element.
        write_print_settings if @embedded && @embedded != 0
      end

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
    #         :line       => { :color => 'blue' }
    #     )
    #
    # The properties that can be set are:
    #
    # ===:values
    #   This is the most important property of a series and must be set
    #   for every chart object. It links the chart with the worksheet data
    #   that it displays. A formula or array ref can be used for the
    #   data range, see below.
    #
    # ===:categories
    #   This sets the chart category labels. The category is more or less
    #   the same as the X-axis. In most chart types the categories property
    #   is optional and the chart will just assume a sequential series
    #   from 1 .. n.
    #
    # ===:name
    #   Set the name for the series. The name is displayed in the chart
    #   legend and in the formula bar. The name property is optional and
    #   if it isn't supplied it will default to Series 1 .. n.
    #
    # ===:line
    #   Set the properties of the series line type such as colour and
    #   width. See the "CHART FORMATTING" section below.
    #
    # ===:border
    #   Set the border properties of the series such as colour and style.
    #   See the "CHART FORMATTING" section below.
    #
    # ===:fill
    #   Set the fill properties of the series such as colour. See the
    #   "CHART FORMATTING"
    #   section below.
    #
    # ==:marker
    #   Set the properties of the series marker such as style and color.
    #   See the "CHART FORMATTING" section below.
    #
    # ===:trendline
    #   Set the properties of the series trendline such as linear,
    #   polynomial and moving average types. See the "CHART FORMATTING"
    #   section below.
    #
    # ===:data_labels
    #   Set data labels for the series. See the "CHART FORMATTING"
    #   section below.
    #
    # ===:invert_if_negative
    #   Invert the fill colour for negative values. Usually only applicable
    #   to column and bar charts.
    #
    # ===:overlap
    # Set the overlap between series in a Bar/Column chart. The range is
    # <tt>+/- 100</tt>. Default is 0.
    #
    #     :overlap => 20
    #
    # Note, it is only necessary to apply this property to one series of the chart.
    #
    # ===:gap
    # Set the gap between series in a Bar/Column chart. The range is
    # <tt>0 to 500</tt>. Default is 150.
    #
    #     :gap => 200,
    #
    # Note, it is only necessary to apply this property to one series of the
    # chart.
    #
    # The categories and values can take either a range formula such
    # as <tt>=Sheet1!$A$2:$A$7</tt> or, more usefully when generating the range
    # programmatically, an array ref with zero indexed row/column values:
    #
    #      [ sheetname, row_start, row_end, col_start, col_end ]
    #
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
    # ==SERIES OPTIONS
    #
    # This section details the following properties of add_series() in more
    # detail:
    #
    #     marker
    #     trendline
    #     y_error_bars
    #     x_error_bars
    #     data_labels
    #     points
    #
    # ===Marker
    #
    # The marker format specifies the properties of the markers used to
    # distinguish series on a chart. In general only Line and Scatter chart
    # types and trendlines use markers.
    #
    # The following properties can be set for marker formats in a chart.
    #
    #     type
    #     size
    #     border
    #     fill
    #
    # The type property sets the type of marker that is used with a series.
    #
    #     chart.add_series(
    #         :values     => '=Sheet1!$B$1:$B$5',
    #         :marker     => { :type => 'diamond' }
    #     )
    #
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
    #
    # The automatic type is a special case which turns on a marker using the
    # default marker style for the particular series number.
    #
    #     chart.add_series(
    #         :values     => '=Sheet1!$B$1:$B$5',
    #         :marker     => { :type => 'automatic' }
    #     )
    #
    # If automatic is on then other marker properties such as size, border or
    # fill cannot be set.
    #
    # The size property sets the size of the marker and is generally used in
    # conjunction with type.
    #
    #     chart.add_series(
    #         :values     => '=Sheet1!$B$1:$B$5',
    #         :marker     => { :type => 'diamond', :size => 7 }
    #     )
    #
    # Nested border and fill properties can also be set for a marker. See the
    # "CHART FORMATTING"
    # section below.
    #
    #     chart.add_series(
    #         :values     => '=Sheet1!$B$1:$B$5',
    #         :marker     => {
    #             :type    => 'square',
    #             :size    => 5,
    #             :border  => { :color => 'red' },
    #             :fill    => { :color => 'yellow' }
    #         }
    #     )
    #
    # ===Trendline
    #
    # A trendline can be added to a chart series to indicate trends in the data
    # such as a moving average or a polynomial fit.
    #
    # The following properties can be set for trendlines in a chart series.
    #
    #     type
    #     order       (for polynomial trends)
    #     period      (for moving average)
    #     forward     (for all except moving average)
    #     backward    (for all except moving average)
    #     name
    #     line
    #
    # The type property sets the type of trendline in the series.
    #
    #     chart.add_series(
    #         :values     => '=Sheet1!$B$1:$B$5',
    #         :trendline  => { :type => 'linear' }
    #     )
    #
    # The available trendline types are:
    #
    #     exponential
    #     linear
    #     log
    #     moving_average
    #     polynomial
    #     power
    #
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
    #
    # A moving_average trendline can also specify the period of the moving
    # average. The default value is 2.
    #
    #     chart.add_series(
    #         :values     => '=Sheet1!$B$1:$B$5',
    #         :trendline  => {
    #             :type   => 'moving_average',
    #             :period => 3,
    #         }
    #     )
    #
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
    #
    # The name property sets an optional name for the trendline that will
    # appear in the chart legend. If it isn't specified the Excel default
    # name will be displayed. This is usually a combination of the
    # trendline type and the series name.
    #
    #     chart.add_series(
    #         :values    => '=Sheet1!$B$1:$B$5',
    #         :trendline => {
    #             :type => 'linear',
    #             :name => 'Interpolated trend'
    #         }
    #     )
    #
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
    #
    # Trendlines cannot be added to series in a stacked chart or pie chart,
    # radar chart or (when implemented) to 3D, surface, or doughnut charts.
    #
    # ===Error Bars
    #
    # Error bars can be added to a chart series to indicate error bounds in the
    # data. The error bars can be vertical y_error_bars (the most common type)
    # or horizontal x_error_bars (for Bar and Scatter charts only).
    #
    # The following properties can be set for error bars in a chart series.
    #
    #     type
    #     value       (for all types except standard error)
    #     direction
    #     end_style
    #     line
    #
    # The type property sets the type of error bars in the series.
    #
    #     chart.add_series(
    #         :values       => '=Sheet1!$B$1:$B$5',
    #         :y_error_bars => { :type => 'standard_error' }
    #     )
    #
    # The available error bars types are available:
    #
    #     fixed
    #     percentage
    #     standard_deviation
    #     standard_error
    #
    # Note, the "custom" error bars type is not supported.
    #
    # All error bar types, except for standard_error must also have a value
    # associated with it for the error bounds:
    #
    #     chart.add_series(
    #         :values       => '=Sheet1!$B$1:$B$5',
    #         :y_error_bars => {
    #             :type  => 'percentage',
    #             :value => 5
    #         }
    #     )
    #
    # The direction property sets the direction of the error bars. It should
    # be one of the following:
    #
    #     plus    # Positive direction only.
    #     minus   # Negative direction only.
    #     both    # Plus and minus directions, The default.
    #
    # The end_style property sets the style of the error bar end cap. The
    # options are 1 (the default) or 0 (for no end cap):
    #
    #     chart.add_series(
    #         :values       => '=Sheet1!$B$1:$B$5',
    #         :y_error_bars => {
    #             :type      => 'fixed',
    #             :value     => 2,
    #             :end_style => 0,
    #             :direction => 'minus'
    #         }
    #     )
    #
    # ===Data Labels
    #
    # Data labels can be added to a chart series to indicate the values of the
    # plotted data points.
    #
    # The following properties can be set for data_labels formats in a chart.
    #
    #     value
    #     category
    #     series_name
    #     position
    #     leader_lines
    #     percentage
    #
    # The value property turns on the Value data label for a series.
    #
    #     chart.add_series(
    #         :values      => '=Sheet1!$B$1:$B$5',
    #         :data_labels => { :value => 1 }
    #     )
    # The category property turns on the Category Name data label for a series.
    #
    #     chart.add_series(
    #         :values      => '=Sheet1!$B$1:$B$5',
    #         :data_labels => { :category => 1 }
    #     )
    #
    # The series_name property turns on the Series Name data label for a series.
    #
    #     chart.add_series(
    #         :values      => '=Sheet1!$B$1:$B$5',
    #         :data_labels => { :series_name => 1 }
    #     )
    #
    # The position property is used to position the data label for a series.
    #
    #     chart.add_series(
    #         :values      => '=Sheet1!$B$1:$B$5',
    #         :data_labels => { :value => 1, :position => 'center' },
    #     )
    #
    # Valid positions are:
    #
    #     center
    #     right
    #     left
    #     top
    #     bottom
    #     above           # Same as top
    #     below           # Same as bottom
    #     inside_end      # Pie chart mainly.
    #     outside_end     # Pie chart mainly.
    #     best_fit        # Pie chart mainly.
    #
    # The percentage property is used to turn on the display of data labels as
    # a Percentage for a series. It is mainly used for pie charts.
    #
    #     chart.add_series(
    #         :values      => '=Sheet1!$B$1:$B$5',
    #         :data_labels => { :percentage => 1 }
    #     )
    #
    # The leader_lines property is used to turn on Leader Lines for the data
    # label for a series. It is mainly used for pie charts.
    #
    #     chart.add_series(
    #         :values      => '=Sheet1!$B$1:$B$5',
    #         :data_labels => { :value => 1, :leader_lines => 1 }
    #     )
    #
    # Note: Even when leader lines are turned on they aren't automatically
    # visible in Excel or Excel::Writer::XLSX. Due to an Excel limitation
    # (or design) leader lines only appear if the data label is moved
    # manually or if the data labels are very close and need to be adjusted
    # automatically.
    #
    # ===Points
    #
    # In general formatting is applied to an entire series in a chart. However,
    # it is occasionally required to format individual points in a series. In
    # particular this is required for Pie charts where each segment is
    # represented by a point.
    #
    # In these cases it is possible to use the points property of add_series():
    #
    #     chart.add_series(
    #         :values => '=Sheet1!$A$1:$A$3',
    #         :points => [
    #             { :fill => { :color => '#FF0000' } },
    #             { :fill => { ?color => '#CC0000' } },
    #             { :fill => { :color => '#990000' } }
    #         ]
    #     )
    #
    # The points property takes an array ref of format options (see the
    # "CHART FORMATTING"
    # section below). To assign default properties to points in a series pass
    # nil values in the array ref:
    #
    #     # Format point 3 of 3 only.
    #     chart.add_series(
    #         :values => '=Sheet1!$A$1:$A$3',
    #         :points => [
    #             nil,
    #             nil,
    #             { :fill => { :color => '#990000' } }
    #         ]
    #     )
    #
    #     # Format the first point only.
    #     chart.add_series(
    #         :values => '=Sheet1!$A$1:$A$3',
    #         :points => [ { :fill => { :color => '#FF0000' } } ]
    #                        )
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
      line = line_properties(params[:line])

      # Allow 'border' as a synonym for 'line' in bar/column style charts.
      line = line_properties(params[:border]) if params[:border]

      # Set the fill properties for the series.
      fill = fill_properties(params[:fill])

      # Set the marker properties for the series.
      marker = marker_properties(params[:marker])

      # Set the trendline properties for the series.
      trendline = trendline_properties(params[:trendline])

      # Set the line smooth property for the series.
      smooth = params[:smooth]

      # Set the error bars properties for the series.
      y_error_bars = error_bars_properties(params[:y_error_bars])
      x_error_bars = error_bars_properties(params[:x_error_bars])

      # Set the point properties for the series.
      points = points_properties(params[:points])

      # Set the labels properties for the series.
      labels = labels_properties(params[:data_labels])

      # Set the "invert if negative" fill property.
      invert_if_neg = params[:invert_if_negative]

      # Set the gap for Bar/Column charts.
      if params[:gap]
        @series_gap = params[:gap]
      end

      # Set the overlap for Bar/Column charts.
      if params[:overlap]
        @series_overlap = params[:overlap]
      end

      # Set the secondary axis properties.
      x2_axis = params[:x2_axis]
      y2_axis = params[:y2_axis]

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
        :_smooth        => smooth,
        :_labels        => labels,
        :_invert_if_neg => invert_if_neg,
        :_x2_axis       => x2_axis,
        :_y2_axis       => y2_axis,
        :_points        => points,
        :_error_bars    => {
          :_x_error_bars => x_error_bars,
          :_y_error_bars => y_error_bars
        }
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
    #     :major_gridlines
    #     :minor_gridlines
    #     :visible
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
    # ==CHART FONTS
    #
    # The following font properties can be set for any chart object that they
    # apply to (and that are supported by WriteXLSX) such as chart titles,
    # axis labels and axis numbering. They correspond to the equivalent
    # Worksheet cell Format object properties. See "FORMAT_METHODS" for more
    # information.
    #
    #     name
    #     size
    #     bold
    #     italic
    #     underline
    #     rotation
    #     color
    #
    # The following explains the available font properties:
    #
    # ===name
    # Set the font name:
    #
    #     chart.set_x_axis( :num_font => { :name => 'Arial' } )
    #
    # ===size
    # Set the font size:
    #
    #     chart.set_x_axis( :num_font => { :name => 'Arial', :size => 10 } )
    #
    # ===bold
    # Set the font bold property, should be 0 or 1:
    #
    #     chart.set_x_axis( :num_font => { :bold => 1 } )
    #
    # ===italic
    # Set the font italic property, should be 0 or 1:
    #
    #     chart.set_x_axis( :num_font => { :italic => 1 } )
    #
    # ===underline
    # Set the font underline property, should be 0 or 1:
    #
    #     chart.set_x_axis( :num_font => { :underline => 1 } )
    #
    # ===rotation
    # See the font rotation in the range -90 to 90:
    #
    #     chart.set_x_axis(:num_font => { :rotation => 45 })
    #
    # This is useful for displaying large axis data such as dates in a more compact format.
    #
    # ===color
    # Set the font color property. Can be a color index, a color name or HTML
    # style RGB colour:
    #
    #     chart.set_x_axis( :num_font => { :color => 'red' } )
    #     chart.set_y_axis( :num_font => { :color => '#92D050' } )
    #
    # Here is an example of Font formatting in a Chart program:
    #
    #     # Format the chart title.
    #     chart.set_title(
    #         :name      => 'Sales Results Chart',
    #         :name_font => {
    #             :name  => 'Calibri',
    #             :color => 'yellow'
    #         }
    #     )
    #
    #     # Format the X-axis.
    #     chart.set_x_axis(
    #         :name      => 'Month',
    #         :name_font => {
    #             :name  => 'Arial',
    #             :color => '#92D050'
    #         },
    #         :num_font => {
    #             :name  => 'Courier New',
    #             :color => '#00B0F0'
    #         }
    #     )
    #
    #     # Format the Y-axis.
    #     chart.set_y_axis(
    #         :name      => 'Sales (1000 units)',
    #         :name_font => {
    #             :name      => 'Century',
    #             :underline => 1,
    #             :color     => 'red'
    #         },
    #         :num_font => {
    #             :bold   => 1,
    #             :italic => 1,
    #             :color  => '#7030A0'
    #         }
    #     )
    def set_x_axis(params = {})
      @x_axis.merge_with_hash(self, params)
    end

    #
    # Set the properties of the Y-axis.
    #
    # The set_y_axis() method is used to set properties of the Y axis.
    # The properties that can be set are the same as for set_x_axis,
    #
    def set_y_axis(params = {})
      @y_axis.merge_with_hash(self, params)
    end

    #
    # Set the properties of the secondary X-axis.
    #
    def set_x2_axis(params = {})
      @x2_axis.merge_with_hash(self, params)
    end

    #
    # Set the properties of the secondary Y-axis.
    #
    def set_y2_axis(params = {})
      @y2_axis.merge_with_hash(self, params)
    end

    #
    # Set the properties of the chart title.
    #
    # The set_title() method is used to set properties of the chart title.
    #
    #     chart.set_title( :name => 'Year End Results' )
    #
    # The properties that can be set are:
    #
    # ===:name
    # Set the name (title) for the chart. The name is displayed above the
    # chart. The name can also be a formula such as +=Sheet1!$A$1+. The name
    # property is optional. The default is to have no chart title.
    #
    # ===:name_font
    # Set the font properties for the chart title. See the "CHART FONTS" section.
    #
    def set_title(params)
      name, name_formula = process_names(params[:name], params[:name_formula])
      data_id = get_data_id(name_formula, params[:data])

      @title_name    = name
      @title_formula = name_formula
      @title_data_id = data_id

      # Set the font properties if present.
      @title_font = convert_font_args(params[:name_font])
    end

    #
    # Set the properties of the chart legend.
    #
    # The set_legend() method is used to set properties of the chart legend.
    #
    #     chart.set_legend( :position => 'none' )
    #
    # The properties that can be set are:
    #
    # ===:position
    # Set the position of the chart legend.
    #
    #     chart.set_legend( :position => 'bottom' )
    #
    # The default legend position is right. The available positions are:
    #
    #     none
    #     top
    #     bottom
    #     left
    #     right
    #     overlay_left
    #     overlay_right
    #
    # ===:delete_series
    #
    # This allows you to remove 1 or more series from the the legend
    # (the series will still display on the chart). This property takes
    # an array ref as an argument and the series are zero indexed:
    #
    #     # Delete/hide series index 0 and 2 from the legend.
    #     chart.set_legend(:delete_series => [0, 2])
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
      # Convert the user defined properties to internal properties.
      @plotarea = area_properties(params)
    end

    #
    # Set the properties of the chart chartarea.
    #
    # The set_chartarea() method is used to set the properties of the chart
    # area.
    #
    #     chart.set_chartarea(
    #         :border => { :none  => 1 },
    #         :fill   => { :color => 'red' }
    #     )
    #
    # The properties that can be set are:
    # ===:border
    # Set the border properties of the chartarea such as colour and style.
    # See the "CHART FORMATTING" section.
    # ===:fill
    # Set the fill properties of the plotarea such as colour. See the
    # "CHART FORMATTING" section.
    #
    def set_chartarea(params)
      # Convert the user defined properties to internal properties.
      @chartarea = area_properties(params)
    end

    #
    # Set on of the 42 built-in Excel chart styles. The default style is 2.
    #
    # The set_style() method is used to set the style of the chart to one
    # of the 42 built-in styles available on the 'Design' tab in Excel:
    #
    #     chart.set_style( 4 )
    #
    # The default style is 2.
    #
    def set_style(style_id = 2)
      style_id = 2 if style_id < 0 || style_id > 42
      @style_id = style_id
    end

    #
    # Set the option for displaying blank data in a chart. The default is 'gap'.
    #
    # The show_blanks_as method controls how blank data is displayed in a chart.
    #
    #    chart.show_blanks_as('span')
    #
    # The available options are:
    #
    #        gap    # Blank data is show as a gap. The default.
    #        zero   # Blank data is displayed as zero.
    #        span   # Blank data is connected with a line.
    #
    def show_blanks_as(option)
      return unless option

      unless [:gap, :zero, :span].include?(option.to_sym)
        raise "Unknown show_blanks_as() option '#{option}'\n"
      end

      @show_blanks = option
    end

    #
    # Display data in hidden rows or columns on the chart.
    #
    def show_hidden_data
      @show_hidden_data = true
    end

    #
    # Set dimensions for scale for the chart.
    #
    # The set_size() method is used to set the dimensions of the chart.
    # The size properties that can be set are:
    #
    #      width
    #      height
    #      x_scale
    #      y_scale
    #      x_offset
    #      y_offset
    #
    # The width and height are in pixels. The default chart width is 480
    # pixels and the default height is 288 pixels. The size of the chart can
    # be modified by setting the width and height or by setting the :x_scale
    # and :y_scale:
    #
    #     chart.set_size( :width => 720, :height => 576 )
    #
    #     # Same as:
    #
    #     chart.set_size( :x_scale => 1.5, :y_scale => 2 )
    #
    # The :x_offset and :y_offset position the top left corner of the chart
    # in the cell that it is inserted into.
    #
    # Note: the :x_scale, :y_scale, :x_offset and :y_offset parameters can also
    # be set via the insert_chart() method:
    #
    #       worksheet.insert_chart( 'E2', chart, 2, 4, 1.5, 2 )
    #
    def set_size(params = {})
      @width    = params[:width]    if params[:width]
      @height   = params[:height]   if params[:height]
      @x_scale  = params[:x_scale]  if params[:x_scale]
      @y_scale  = params[:y_scale]  if params[:y_scale]
      @x_offset = params[:x_offset] if params[:x_offset]
      @y_offset = params[:y_offset] if params[:y_offset]
    end

    # Backward compatibility with poorly chosen method name.
    alias :size :set_size

    #
    # The set_table method adds a data table below the horizontal axis with the
    # data used to plot the chart.
    #
    #    chart.set_table
    #
    # The available options, with default values are:
    #
    #    :vertical   => true    # Display vertical lines in the table.
    #    :horizontal => true    # Display horizontal lines in the table.
    #    :outline    => true    # Display an outline in the table.
    #    :show_keys  => false   # Show the legend keys with the table data.
    #
    # The data table can only be shown with Bar, Column, Line, Area and Stock charts.
    #
    def set_table(params = {})
      @table = Table.new(params)
    end

    #
    # Set properties for the chart up-down bars.
    #
    # The set_up_down_bars() method adds Up-Down bars to Line charts to
    # indicate the difference between the first and last data series.
    #
    #     chart.set_up_down_bars
    # It is possible to format the up and down bars to add fill and border
    # properties if required. See the "CHART FORMATTING" section below.
    #
    #     chart.set_up_down_bars(
    #         :up   => { :fill => { :color => 'green' } },
    #         :down => { :fill => { :color => 'red' } }
    #     )
    # Up-down bars can only be applied to Line charts and to Stock charts
    # (by default).
    #
    def set_up_down_bars(params = {})
      # Map border to line.
      [:up, :down].each do |up_down|
        if params[up_down]
          params[up_down][:line] = params[up_down][:border] if params[up_down][:border]
        else
          params[up_down] = {}
        end
      end

      # Set the up and down bar properties.
      @up_down_bars = {
        :_up => {
          :_line => line_properties(params[:up][:line]),
          :_fill => line_properties(params[:up][:fill])
        },
        :_down => {
          :_line => line_properties(params[:down][:line]),
          :_fill => line_properties(params[:down][:fill])
        }
      }
    end

    #
    # Set properties for the chart drop lines.
    #
    # The set_drop_lines() method adds Drop Lines to charts to show the
    # Category value of points in the data.
    #
    #     chart.set_drop_lines
    #
    # It is possible to format the Drop Line line properties if required.
    # See the "CHART FORMATTING" section below.
    #
    #     chart.set_drop_lines(:line => { :color => 'red', :dash_type => 'square_dot' } )
    #
    #     Drop Lines are only available in Line, Area and Stock charts.
    #
    def set_drop_lines(params = {})
      # Set the drop line properties.
      line = line_properties(params[:line])

      @drop_lines = { :_line => line }
    end

    #
    # Set properties for the chart high-low lines.
    #
    # The set_high_low_lines() method adds High-Low lines to charts to show
    # the maximum and minimum values of points in a Category.
    #
    #     chart.set_high_low_lines
    #
    # It is possible to format the High-Low Line line properties if required.
    # See the "CHART FORMATTING" section below.
    #
    #     chart.set_high_low_lines( :line => { :color => 'red' } )
    #
    #     High-Low Lines are only available in Line and Stock charts.
    #
    def set_high_low_lines(params = {})
      # Set the drop line properties.
      line = line_properties(params[:line])

      @hi_low_lines = { :_line => line }
    end

    #
    # Setup the default configuration data for an embedded chart.
    #
    def set_embedded_config_data
      @embedded = 1
    end

    #
    # Write the <c:barChart> element.
    #
    def write_bar_chart(params)   # :nodoc:
      if ptrue?(params[:primary_axes])
        series = get_primary_axes_series
      else
        series = get_secondary_axes_series
      end
      return if series.empty?

      subtype = @subtype
      subtype = 'percentStacked' if subtype == 'percent_stacked'

      # Set a default overlap for stacked charts.
      if @subtype =~ /stacked/
        @series_overlap = 100 unless @series_overlap
      end

      @writer.tag_elements('c:barChart') do
        # Write the c:barDir element.
        write_bar_dir
        # Write the c:grouping element.
        write_grouping(subtype)
        # Write the c:ser elements.
        series.each {|s| write_ser(s)}

        # write the c:marker element.
        write_marker_value

        # Write the c:gapWidth element.
        write_gap_width(@series_gap)

        # write the c:overlap element.
        write_overlap(@series_overlap)

        # Write the c:axId elements
        write_axis_ids(params)
      end
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
    # Convert user defined font values into private hash values.
    #
    def convert_font_args(params)
      return unless params
      font = {
        :_name         => params[:name],
        :_color        => params[:color],
        :_size         => params[:size],
        :_bold         => params[:bold],
        :_italic       => params[:italic],
        :_underline    => params[:underline],
        :_pitch_family => params[:pitch_family],
        :_charset      => params[:charset],
        :_baseline     => params[:baseline] || 0,
        :_rotation     => params[:rotation]
      }

      # Convert font size units.
      font[:_size] *= 100 if font[:_size] && font[:_size] != 0

      # Convert rotation into 60,000ths of a degree.
      if ptrue?(font[:_rotation])
        font[:_rotation] = 60_000 * font[:_rotation].to_i
      end

      font
    end

    #
    # Convert user defined line properties to the structure required internally.
    #
    def line_properties(line) # :nodoc:
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
        line[:dash_type] = value_or_raise(dash_types, dash_type, 'dash type')
      end

      line[:_defined] = 1

      line
    end

    private

    #
    # retun primary/secondary series by :primary_axes flag
    #
    def axes_series(params)
      if params[:primary_axes] != 0
        primary_axes_series
      else
        secondary_axes_series
      end
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
    # Convert the user specified colour index or string to a rgb colour.
    #
    def get_color(color) # :nodoc:
      # Convert a HTML style #RRGGBB color.
      if color and color =~ /^#[0-9a-fA-F]{6}$/
        color = color.sub(/^#/, '')
        return color.upcase
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
    # Get the Spreadsheet::WriteExcel line pattern for backward compatibility.
    #
    def get_swe_line_pattern(val)
      value   = val.downcase
      default = 'solid'

      patterns = {
        0              => 'solid',
        1              => 'dash',
        2              => 'dot',
        3              => 'dash_dot',
        4              => 'long_dash_dot_dot',
        5              => 'none',
        6              => 'solid',
        7              => 'solid',
        8              => 'solid',
        'solid'        => 'solid',
        'dash'         => 'dash',
        'dot'          => 'dot',
        'dash-dot'     => 'dash_dot',
        'dash-dot-dot' => 'long_dash_dot_dot',
        'none'         => 'none',
        'dark-gray'    => 'solid',
        'medium-gray'  => 'solid',
        'light-gray'   => 'solid'
      }

      patterns[value] || default
    end

    #
    # Get the Spreadsheet::WriteExcel line weight for backward compatibility.
    #
    def get_swe_line_weight(val)
      value   = val.downcase
      default = 1

      weights = {
        1          => 0.25,
        2          => 1,
        3          => 2,
        4          => 3,
        'hairline' => 0.25,
        'narrow'   => 1,
        'medium'   => 2,
        'wide'     => 3
      }

      weights[value] || default
    end

    #
    # Convert user defined fill properties to the structure required internally.
    #
    def fill_properties(fill) # :nodoc:
      return { :_defined => 0 } unless fill

      fill[:_defined] = 1

      fill
    end

    #
    # Convert user defined marker properties to the structure required internally.
    #
    def marker_properties(marker) # :nodoc:
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
      marker_type = marker[:type]

      if marker_type
        marker[:automatic] = 1 if marker_type == 'automatic'
        marker[:type] = value_or_raise(types, marker_type, 'maker type')
      end

      # Set the line properties for the marker..
      line = line_properties(marker[:line])

      # Allow 'border' as a synonym for 'line'.
      line = line_properties(marker[:border]) if marker[:border]

      # Set the fill properties for the marker.
      fill = fill_properties(marker[:fill])

      marker[:_line] = line
      marker[:_fill] = fill

      marker
    end

    #
    # Convert user defined trendline properties to the structure required internally.
    #
    def trendline_properties(trendline) # :nodoc:
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
      trend_type = trendline[:type]

      trendline[:type] = value_or_raise(types, trend_type, 'trendline type')

      # Set the line properties for the trendline..
      line = line_properties(trendline[:line])

      # Allow 'border' as a synonym for 'line'.
      line = line_properties(trendline[:border]) if trendline[:border]

      # Set the fill properties for the trendline.
      fill = fill_properties(trendline[:fill])

      trendline[:_line] = line
      trendline[:_fill] = fill

      return trendline
    end

    #
    # Convert user defined error bars properties to structure required
    # internally.
    #
    def error_bars_properties(params = {})
      return if !ptrue?(params) || params.empty?

      # Default values.
      error_bars = {
        :_type      => 'fixedVal',
        :_value     => 1,
        :_endcap    => 1,
        :_direction => 'both'
      }

      types = {
        :fixed              => 'fixedVal',
        :percentage         => 'percentage',
        :standard_deviation => 'stdDev',
        :standard_error     => 'stdErr'
      }

      # Check the error bars type.
      error_type = params[:type].to_sym

      if types.key?(error_type)
        error_bars[:_type] = types[error_type]
      else
        raise "Unknown error bars type '#{error_type}'\n"
      end

      # Set the value for error types that require it.
      if params.key?(:value)
        error_bars[:_value] = params[:value]
      end

      # Set the end-cap style.
      if params.key?(:end_style)
        error_bars[:_endcap] = params[:end_style]
      end

      # Set the error bar direction.
      if params.key?(:direction)
        if params[:direction] == 'minus'
          error_bars[:_direction] = 'minus'
        elsif params[:direction] == 'plus'
          error_bars[:_direction] = 'plus'
        else
          # Default to 'both'
        end
      end

      # Set the line properties for the error bars.
      error_bars[:_line] = line_properties(params[:line])

      error_bars
    end

    #
    # Convert user defined labels properties to the structure required internally.
    #
    def labels_properties(labels) # :nodoc:
      return nil unless labels

      position = labels[:position]
      if position.nil? || position.empty?
        labels.delete(:position)
      else
        # Map user defined label positions to Excel positions.
        positions = {
          :center      => 'ctr',
          :right       => 'r',
          :left        => 'l',
          :top         => 't',
          :above       => 't',
          :bottom      => 'b',
          :below       => 'b',
          :inside_end  => 'inEnd',
          :outside_end => 'outEnd',
          :best_fit    => 'bestFit'
        }

        labels[:position] = value_or_raise(positions, position, 'label position')
      end

      labels
    end

    #
    # Convert user defined area properties to the structure required internally.
    #
    def area_properties(arg)  # :nodoc:
      area = {}

      # Map deprecated Spreadsheet::WriteExcel fill colour.
      arg[:fill] = { :color => arg[:color] } if arg[:color]

      # Map deprecated Spreadsheet::WriteExcel line_weight.
      if arg[:line_weight]
        width = get_swe_line_weight(arg[:line_weight])
        arg[:border] = { :width => width }
      end

      # Map deprecated Spreadsheet::WriteExcel line_pattern.
      if arg[:line_pattern]
        pattern = get_swe_line_pattern(arg[:line_pattern])
        if pattern == 'none'
          arg[:border] = { :none => 1 }
        else
          arg[:border][:dash_type] = pattern
        end
      end

      # Map deprecated Spreadsheet::WriteExcel line colour.
      arg[:border][:color] = arg[:line_color] if arg[:line_color]

      # Handle Excel::Writer::XLSX style properties.

      # Set the line properties for the chartarea.
      line = line_properties(arg[:line])

      # Allow 'border' as a synonym for 'line'.
      line = line_properties(arg[:border]) if (arg[:border])

      # Set the fill properties for the chartarea.
      fill = fill_properties(arg[:fill])

      area[:_line] = line
      area[:_fill] = fill

      return area
    end

    #
    # Convert user defined points properties to structure required internally.
    #
    def points_properties(user_points = nil)
      return unless user_points

      points = []
      user_points.each do |user_point|
        if user_point
          # Set the lline properties for the point.
          line = line_properties(user_point[:line])

          # Allow 'border' as a synonym for 'line'.
          if user_point[:border]
            line = line_properties(user_point[:border])
          end

          # Set the fill properties for the chartarea.
          fill = fill_properties(user_point[:fill])

          point = {}
          point[:_line] = line
          point[:_fill] = fill
        end
        points << point
      end
      points
    end

    def value_or_raise(hash, key, msg)
      raise "Unknown #{msg} '#{key}'" unless hash[key.to_sym]
      hash[key.to_sym]
    end

    #
    # Returns series which use the primary axes.
    #
    def get_primary_axes_series
      @series.reject {|s| s[:_y2_axis]}
    end
    alias :primary_axes_series :get_primary_axes_series

    #
    # Returns series which use the secondary axes.
    #
    def get_secondary_axes_series
      @series.select {|s| s[:_y2_axis]}
    end
    alias :secondary_axes_series :get_secondary_axes_series

    #
    # Add a unique ids for primary or secondary axis.
    #
    def add_axis_ids(params) # :nodoc:
      if ptrue?(params[:primary_axes])
        @axis_ids  += ids
      else
        @axis2_ids += ids
      end
    end

    def ids
      chart_id   = 1 + @id
      axis_count = 1 + @axis2_ids.size + @axis_ids.size

      id1 = sprintf('5%03d%04d', chart_id, axis_count)
      id2 = sprintf('5%03d%04d', chart_id, axis_count + 1)

      [id1, id2]
    end

    #
    # Get the font style attributes from a font hash.
    #
    def get_font_style_attributes(font)
      return [] unless font

      attributes = []
      attributes << 'sz' << font[:_size]      if ptrue?(font[:_size])
      attributes << 'b'  << font[:_bold]      if font[:_bold]
      attributes << 'i'  << font[:_italic]    if font[:_italic]
      attributes << 'u'  << 'sng'             if font[:_underline]

      attributes << 'baseline' << font[:_baseline]
      attributes
    end

    #
    # Get the font latin attributes from a font hash.
    #
    def get_font_latin_attributes(font)
      return [] unless font

      attributes = []
      attributes << 'typeface' << font[:_name] if ptrue?(font[:_name])
      attributes << 'pitchFamily' << font[:_pitch_family] if font[:_pitch_family]
      attributes << 'charset' << font[:_charset] if font[:_charset]

      attributes
    end
    #
    # Setup the default properties for a chart.
    #
    def set_default_properties # :nodoc:
      # Set the default axis properties.
      @x_axis.defaults = {
        :num_format      => 'General',
        :major_gridlines => { :visible => 0 }
      }

      @y_axis.defaults = {
        :num_format      => 'General',
        :major_gridlines => { :visible => 1 }
      }

      @x2_axis.defaults = {
        :num_format     => 'General',
        :label_position => 'none',
        :crossing       => 'max',
        :visible        => 0
      }

      @y2_axis.defaults = {
        :num_format      => 'General',
        :major_gridlines => { :visible => 0 },
        :position        => 'right',
        :visible         => 1
      }

      set_x_axis
      set_y_axis

      set_x2_axis
      set_y2_axis
    end

    def default_chartarea_property
      {
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
    end

    def default_chartarea_property_for_embedded
      default_chartarea_property.
        merge(
              :_visible => 1,
              :_area_pattern => 0x0001,
              :_area_options => 0x0001,
              :_line_pattern => 0x0000,
              :_line_weight  => 0x0000,
              :_line_options => 0x0009
              )
    end

    def default_plotarea_property
      {
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
      @writer.tag_elements('c:chartSpace', chart_space_attributes) do
        yield
      end
    end

    # for <c:chartSpace> element.
    def chart_space_attributes # :nodoc:
      schema  = 'http://schemas.openxmlformats.org/'
      [
       'xmlns:c', "#{schema}drawingml/2006/chart",
       'xmlns:a', "#{schema}drawingml/2006/main",
       'xmlns:r', "#{schema}officeDocument/2006/relationships"
      ]
    end

    #
    # Write the <c:lang> element.
    #
    def write_lang # :nodoc:
      @writer.empty_tag('c:lang', ['val', 'en-US'])
    end

    #
    # Write the <c:style> element.
    #
    def write_style # :nodoc:
      return if @style_id == 2

      @writer.empty_tag('c:style', ['val', @style_id])
    end

    #
    # Write the <c:chart> element.
    #
    def write_chart # :nodoc:
      @writer.tag_elements('c:chart') do
        # Write the chart title elements.
        if title = @title_formula
          write_title_formula(title, @title_data_id, nil, @title_font)
        elsif title = @title_name
          write_title_rich(title, nil, @title_font)
        end

        # Write the c:plotArea element.
        write_plot_area
        # Write the c:legend element.
        write_legend
        # Write the c:plotVisOnly element.
        write_plot_vis_only

        # Write the c:dispBlanksAs element.
        write_disp_blanks_as
      end
    end

    #
    # Write the <c:dispBlanksAs> element.
    #
    def write_disp_blanks_as
      return if @show_blanks == 'gap'

      @writer.empty_tag('c:dispBlanksAs', ['val', @show_blanks])
    end

    #
    # Write the <c:plotArea> element.
    #

    def write_plot_area   # :nodoc:
      write_plot_area_base
    end

    def write_plot_area_base(type = nil) # :nodoc:
      @writer.tag_elements('c:plotArea') do
        # Write the c:layout element.
        write_layout
        # Write the subclass chart type elements for primary and secondary axes.
        write_chart_type(:primary_axes => 1)
        write_chart_type(:primary_axes => 0)

        # Write the c:catAx elements for series using primary axes.
        params = {
          :x_axis   => @x_axis,
          :y_axis   => @y_axis,
          :axis_ids => @axis_ids
        }
        write_cat_or_date_axis(params, type)
        write_val_axis(params)

        # Write c:valAx and c:catAx elements for series using secondary axes.
        params = {
          :x_axis   => @x2_axis,
          :y_axis   => @y2_axis,
          :axis_ids => @axis2_ids
        }
        write_val_axis(params)
        write_cat_or_date_axis(params, type)

        # Write the c:dTable element.
        write_d_table

        # Write the c:spPr element for the plotarea formatting.
        write_sp_pr(@plotarea)
      end
    end

    def write_cat_or_date_axis(params, type)
      if type == :stock
        write_date_axis(params)
      else
        write_cat_axis(params)
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
      @writer.empty_tag('c:grouping', ['val', val])
    end

    #
    # Write the series elements.
    #
    def write_series(series) # :nodoc:
      write_ser(series)
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
    def write_ser(series) # :nodoc:
      index = @series_index
      @series_index += 1

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
        # Write the c:dPt element.
        write_d_pt(series[:_points])
        # Write the c:dLbls element.
        write_d_lbls(series[:_labels])
        # Write the c:trendline element.
        write_trendline(series[:_trendline])
        # Write the c:errBars element.
        write_error_bars(series[:_error_bars])
        # Write the c:cat element.
        write_cat(series)
        # Write the c:val element.
        write_val(series)
        # Write the c:smooth element.
        write_c_smooth(series[:_smooth]) if ptrue?(@smooth_allowed)
      end
    end

    #
    # Write the <c:idx> element.
    #
    def write_idx(val) # :nodoc:
      @writer.empty_tag('c:idx', ['val', val])
    end

    #
    # Write the <c:order> element.
    #
    def write_order(val) # :nodoc:
      @writer.empty_tag('c:order', ['val', val])
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

      @writer.tag_elements('c:cat') do
        # Check the type of cached data.
        type = get_data_type(data)
        if type == 'str'
          @cat_has_num_fmt = false
          # Write the c:strRef element.
          write_str_ref(formula, data, type)
        else
          @cat_has_num_fmt = true
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
        # Unlike Cat axes data should only be numeric.

        # Write the c:numRef element.
        write_num_ref(formula, data, 'num')
      end
    end

    #
    # Write the <c:numRef> or <c:strRef> element.
    #
    def write_num_or_str_ref(tag, formula, data, type) # :nodoc:
      @writer.tag_elements(tag) do
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
    # Write the <c:numRef> element.
    #
    def write_num_ref(formula, data, type) # :nodoc:
      write_num_or_str_ref('c:numRef', formula, data, type)
    end

    #
    # Write the <c:strRef> element.
    #
    def write_str_ref(formula, data, type) # :nodoc:
      write_num_or_str_ref('c:strRef', formula, data, type)
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
    # Write the <c:axId> elements for the primary or secondary axes.
    #
    def write_axis_ids(params)
      # Generate the axis ids.
      add_axis_ids(params)

      if params[:primary_axes] != 0
        # Write the axis ids for the primary axes.
        write_axis_id(@axis_ids[0])
        write_axis_id(@axis_ids[1])
      else
        # Write the axis ids for the secondary axes.
        write_axis_id(@axis2_ids[0])
        write_axis_id(@axis2_ids[1])
      end
    end

    #
    # Write the <c:axId> element.
    #
    def write_axis_id(val) # :nodoc:
      @writer.empty_tag('c:axId', ['val', val])
    end

    #
    # Write the <c:catAx> element. Usually the X axis.
    #
    def write_cat_axis(params) # :nodoc:
      x_axis   = params[:x_axis]
      y_axis   = params[:y_axis]
      axis_ids = params[:axis_ids]

      # if there are no axis_ids then we don't need to write this element
      return unless axis_ids
      return if axis_ids.empty?

      position = @cat_axis_position
      horiz    = @horiz_cat_axis

      # Overwrite the default axis position with a user supplied value.
      position = x_axis.position || position

      @writer.tag_elements('c:catAx') do
        write_axis_id(axis_ids[0])
        # Write the c:scaling element.
        write_scaling(x_axis.reverse)

        write_delete(1) unless ptrue?(x_axis.visible)

        # Write the c:axPos element.
        write_axis_pos(position, y_axis.reverse)

        # Write the c:majorGridlines element.
        write_major_gridlines(x_axis.major_gridlines)

        # Write the c:minorGridlines element.
        write_minor_gridlines(x_axis.minor_gridlines)

        # Write the axis title elements.
        if title = x_axis.formula
          write_title_formula(title, @x_axis.data_id, horiz, @x_axis.name_font)
        elsif title = x_axis.name
          write_title_rich(title, horiz, x_axis.name_font)
        end

        # Write the c:numFmt element.
        write_cat_number_format(x_axis)

        # Write the c:majorTickMark element.
        write_major_tick_mark(x_axis.major_tick_mark)

        # Write the c:tickLblPos element.
        write_tick_label_pos(x_axis.label_position)

        # Write the axis font elements.
        write_axis_font(x_axis.num_font)

        # Write the c:crossAx element.
        write_cross_axis(axis_ids[1])

        if @show_crosses || ptrue?(x_axis.visible)
          write_crossing(y_axis.crossing)
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
    def write_val_axis(params) # :nodoc:
      x_axis   = params[:x_axis]
      y_axis   = params[:y_axis]
      axis_ids = params[:axis_ids]
      position = params[:position] || @val_axis_position
      horiz    = @horiz_val_axis

      return unless axis_ids && !axis_ids.empty?

      # OVerwrite the default axis position with a user supplied value.
      position = y_axis.position || position

      @writer.tag_elements('c:valAx') do
        write_axis_id(axis_ids[1])

        # Write the c:scaling element.
        write_scaling_with_param(y_axis)

        write_delete(1) unless ptrue?(y_axis.visible)

        # Write the c:axPos element.
        write_axis_pos(position, x_axis.reverse)

        # Write the c:majorGridlines element.
        write_major_gridlines(y_axis.major_gridlines)

        # Write the c:minorGridlines element.
        write_minor_gridlines(y_axis.minor_gridlines)

        # Write the axis title elements.
        if title = y_axis.formula
          write_title_formula(title, y_axis.data_id, horiz, y_axis.name_font)
        elsif title = y_axis.name
          write_title_rich(title, horiz, y_axis.name_font)
        end

        # Write the c:numberFormat element.
        write_number_format(y_axis)

        # Write the c:majorTickMark element.
        write_major_tick_mark(y_axis.major_tick_mark)

        # Write the tickLblPos element.
        write_tick_label_pos(y_axis.label_position)

        # Write the axis font elements.
        write_axis_font(y_axis.num_font)

        # Write the c:crossAx element.
        write_cross_axis(axis_ids[0])

        write_crossing(x_axis.crossing)

        # Write the c:crossBetween element.
        write_cross_between

        # Write the c:majorUnit element.
        write_c_major_unit(y_axis.major_unit)

        # Write the c:minorUnit element.
        write_c_minor_unit(y_axis.minor_unit)
      end
    end

    #
    # Write the <c:valAx> element.
    # This is for the second valAx in scatter plots.
    #
    # Usually the X axis.
    #
    def write_cat_val_axis(params) # :nodoc:
      x_axis   = params[:x_axis]
      y_axis   = params[:y_axis]
      axis_ids = params[:axis_ids]
      position = params[:position] || @val_axis_position
      horiz    = @horiz_val_axis

      return unless axis_ids && !axis_ids.empty?

      # Overwrite the default axis position with a user supplied value.
      position = x_axis.position || position

      @writer.tag_elements('c:valAx') do
        write_axis_id(axis_ids[0])

        # Write the c:scaling element.
        write_scaling_with_param(x_axis)

        write_delete(1) unless ptrue?(x_axis.visible)

        # Write the c:axPos element.
        write_axis_pos(position, y_axis.reverse)

        # Write the c:majorGridlines element.
        write_major_gridlines(x_axis.major_gridlines)

        # Write the c:minorGridlines element.
        write_minor_gridlines(x_axis.minor_gridlines)

        # Write the axis title elements.
        if title = x_axis.formula
          write_title_formula(title, y_axis.data_id, horiz, x_axis.name_font)
        elsif title = x_axis.name
          write_title_rich(title, horiz, x_axis.name_font)
        end

        # Write the c:numberFormat element.
        write_number_format(x_axis)

        # Write the c:majorTickMark element.
        write_major_tick_mark(x_axis.major_tick_mark)

        # Write the c:tickLblPos element.
        write_tick_label_pos(x_axis.label_position)

        # Write the axis font elements.
        write_axis_font(x_axis.num_font)

        # Write the c:crossAx element.
        write_cross_axis(axis_ids[1])

        write_crossing(y_axis.crossing)

        # Write the c:crossBetween element.
        write_cross_between

        # Write the c:majorUnit element.
        write_c_major_unit(x_axis.major_unit)

        # Write the c:minorunit element.
        write_c_minor_unit(x_axis.minor_unit)
      end
    end

    #
    # Write the <c:dateAx> element. Usually the X axis.
    #
    def write_date_axis(params)  # :nodoc:
      x_axis    = params[:x_axis]
      y_axis    = params[:y_axis]
      axis_ids  = params[:axis_ids]

      return unless axis_ids && !axis_ids.empty?

      position  = @cat_axis_position

      # Overwrite the default axis position with a user supplied value.
      position = x_axis.position || position

      @writer.tag_elements('c:dateAx') do
        write_axis_id(axis_ids[0])
        # Write the c:scaling element.
        write_scaling_with_param(x_axis)

        write_delete(1) unless ptrue?(x_axis.visible)

        # Write the c:axPos element.
        write_axis_pos(position, y_axis.reverse)

        # Write the c:majorGridlines element.
        write_major_gridlines(x_axis.major_gridlines)

        # Write the c:minorGridlines element.
        write_minor_gridlines(x_axis.minor_gridlines)

        # Write the axis title elements.
        if title = x_axis.formula
          write_title_formula(title, x_axis.data_id, nil, x_axis.name_font)
        elsif title = x_axis.name
          write_title_rich(title, nil, x_axis.name_font)
        end
        # Write the c:numFmt element.
        write_number_format(x_axis)
        # Write the c:majorTickMark element.
        write_major_tick_mark(x_axis.major_tick_mark)

        # Write the c:tickLblPos element.
        write_tick_label_pos(x_axis.label_position)
        # Write the font elements.
        write_axis_font(x_axis.num_font)
        # Write the c:crossAx element.
        write_cross_axis(axis_ids[1])

        if @show_crosses || ptrue?(x_axis.visible)
          write_crossing(y_axis.crossing)
        end

        # Write the c:auto element.
        write_auto(1)
        # Write the c:labelOffset element.
        write_label_offset(100)
        # Write the c:majorUnit element.
        write_c_major_unit(x_axis.major_unit)
        # Write the c:majorTimeUnit element.
        if !x_axis.major_unit.nil?
          write_c_major_time_unit(x_axis.major_unit_type)
        end
        # Write the c:minorUnit element.
        write_c_minor_unit(x_axis.minor_unit)
        # Write the c:minorTimeUnit element.
        if !x_axis.minor_unit.nil?
          write_c_minor_time_unit(x_axis.minor_unit_type)
        end
      end
    end

    def write_crossing(crossing)
      # Note, the category crossing comes from the value axis.
      if nil_or_max?(crossing)
        # Write the c:crosses element.
        write_crosses(crossing)
      else
        # Write the c:crossesAt element.
        write_c_crosses_at(crossing)
      end
    end

    def write_scaling_with_param(param)
      write_scaling(
                    param.reverse,
                    param.min,
                    param.max,
                    param.log_base
                    )
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
    # Write the <c:logBase> element.
    #
    def write_c_log_base(val) # :nodoc:
      return unless ptrue?(val)

      @writer.empty_tag('c:logBase', ['val', val])
    end

    #
    # Write the <c:orientation> element.
    #
    def write_orientation(reverse = nil) # :nodoc:
      val     = ptrue?(reverse) ? 'maxMin' : 'minMax'

      @writer.empty_tag('c:orientation', ['val', val])
    end

    #
    # Write the <c:max> element.
    #
    def write_c_max(max = nil) # :nodoc:
      return if max.nil?

      @writer.empty_tag('c:max', ['val', max])
    end

    #
    # Write the <c:min> element.
    #
    def write_c_min(min = nil) # :nodoc:
      return if min.nil?

      @writer.empty_tag('c:min', ['val', min])
    end

    #
    # Write the <c:axPos> element.
    #
    def write_axis_pos(val, reverse = false) # :nodoc:
      if reverse
        val = 'r' if val == 'l'
        val = 't' if val == 'b'
      end

      @writer.empty_tag('c:axPos', ['val', val])
    end

    #
    # Write the <c:numberFormat> element. Note: It is assumed that if a user
    # defined number format is supplied (i.e., non-default) then the sourceLinked
    # attribute is 0. The user can override this if required.
    #

    def write_number_format(axis) # :nodoc:
      axis.write_number_format(@writer)
    end

    #
    # Write the <c:numFmt> element. Special case handler for category axes which
    # don't always have a number format.
    #
    def write_cat_number_format(axis)
      axis.write_cat_number_format(@writer, @cat_has_num_fmt)
    end

    #
    # Write the <c:majorTickMark> element.
    #
    def write_major_tick_mark(val)
      return unless ptrue?(val)

      @writer.empty_tag('c:majorTickMark', ['val', val])
    end

    #
    # Write the <c:tickLblPos> element.
    #
    def write_tick_label_pos(val) # :nodoc:
      val ||= 'nextTo'
      val = 'nextTo' if val == 'next_to'

      @writer.empty_tag('c:tickLblPos', ['val', val])
    end

    #
    # Write the <c:crossAx> element.
    #
    def write_cross_axis(val = 'autoZero') # :nodoc:
      @writer.empty_tag('c:crossAx', ['val', val])
    end

    #
    # Write the <c:crosses> element.
    #
    def write_crosses(val) # :nodoc:
      val ||= 'autoZero'

      @writer.empty_tag('c:crosses', ['val', val])
    end

    #
    # Write the <c:crossesAt> element.
    #
    def write_c_crosses_at(val) # :nodoc:
      @writer.empty_tag('c:crossesAt', ['val', val])
    end

    #
    # Write the <c:auto> element.
    #
    def write_auto(val) # :nodoc:
      @writer.empty_tag('c:auto', ['val', val])
    end

    #
    # Write the <c:labelAlign> element.
    #
    def write_label_align(val) # :nodoc:
      @writer.empty_tag('c:lblAlgn', ['val', val])
    end

    #
    # Write the <c:labelOffset> element.
    #
    def write_label_offset(val) # :nodoc:
      @writer.empty_tag('c:lblOffset', ['val', val])
    end

    #
    # Write the <c:majorGridlines> element.
    #
    def write_major_gridlines(gridlines) # :nodoc:
      write_gridlines_base('c:majorGridlines', gridlines)
    end

    #
    # Write the <c:minorGridlines> element.
    #
    def write_minor_gridlines(gridlines)  # :nodoc:
      write_gridlines_base('c:minorGridlines', gridlines)
    end

    def write_gridlines_base(tag, gridlines)  # :nodoc:
      return if gridlines.respond_to?(:[]) and !ptrue?(gridlines[:_visible])
      write_lines_base(tag, gridlines)
    end

    #
    # Write the <c:crossBetween> element.
    #
    def write_cross_between # :nodoc:
      val  = @cross_between || 'between'

      @writer.empty_tag('c:crossBetween', ['val', val])
    end

    #
    # Write the <c:majorUnit> element.
    #
    def write_c_major_unit(val = nil) # :nodoc:
      return unless val

      @writer.empty_tag('c:majorUnit', ['val', val])
    end

    #
    # Write the <c:minorUnit> element.
    #
    def write_c_minor_unit(val = nil) # :nodoc:
      return unless val

      @writer.empty_tag('c:minorUnit', ['val', val])
    end

    #
    # Write the <c:majorTimeUnit> element.
    #
    def write_c_major_time_unit(val) # :nodoc:
      val ||= 'days'

      @writer.empty_tag('c:majorTimeUnit', ['val', val])
    end

    #
    # Write the <c:minorTimeUnit> element.
    #
    def write_c_minor_time_unit(val) # :nodoc:
      val ||= 'days'

      @writer.empty_tag('c:minorTimeUnit', ['val', val])
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
      @writer.empty_tag('c:legendPos', ['val', val])
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
      @writer.empty_tag('c:overlay', ['val', 1])
    end

    #
    # Write the <c:plotVisOnly> element.
    #
    def write_plot_vis_only # :nodoc:
      val  = 1

      # Ignore this element if we are plotting hidden data.
      return if @show_hidden_data

      @writer.empty_tag('c:plotVisOnly', ['val', val])
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
    def write_title_rich(title, horiz = nil, font = nil) # :nodoc:
      @writer.tag_elements('c:title') do
        # Write the c:tx element.
        write_tx_rich(title, horiz, font)
        # Write the c:layout element.
        write_layout
      end
    end

    #
    # Write the <c:title> element for a rich string.
    #
    def write_title_formula(title, data_id, horiz = nil, font = nil) # :nodoc:
      @writer.tag_elements('c:title') do
        # Write the c:tx element.
        write_tx_formula(title, data_id)
        # Write the c:layout element.
        write_layout
        # Write the c:txPr element.
        write_tx_pr(horiz, font)
      end
    end

    #
    # Write the <c:tx> element.
    #
    def write_tx_rich(title, horiz, font = nil) # :nodoc:
      @writer.tag_elements('c:tx') { write_rich(title, horiz, font) }
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
    def write_rich(title, horiz, font) # :nodoc:
      @writer.tag_elements('c:rich') do
        # Write the a:bodyPr element.
        write_a_body_pr(horiz)
        # Write the a:lstStyle element.
        write_a_lst_style
        # Write the a:p element.
        write_a_p_rich(title, font)
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

      attributes = [] unless ptrue?(horiz)

      @writer.empty_tag('a:bodyPr', attributes)
    end

    #
    # Write the <a:bodyPr> element for axis fonts.
    #
    def write_axis_body_pr(rot = nil, vert = nil)
      attributes = []

      attributes << 'rot'  << rot  if rot
      attributes << 'vert' << vert if vert

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
    def write_a_p_rich(title, font) # :nodoc:
      @writer.tag_elements('a:p') do
        # Write the a:pPr element.
        write_a_p_pr_rich(font)
        # Write the a:r element.
        write_a_r(title, font)
      end
    end

    #
    # Write the <a:p> element for formula titles.
    #
    def write_a_p_formula(font = nil) # :nodoc:
      @writer.tag_elements('a:p') do
        # Write the a:pPr element.
        write_a_p_pr_formula(font)
        # Write the a:endParaRPr element.
        write_a_end_para_rpr
      end
    end

    #
    # Write the <a:pPr> element for rich string titles.
    #
    def write_a_p_pr_rich(font) # :nodoc:
      @writer.tag_elements('a:pPr') { write_a_def_rpr(font) }
    end

    #
    # Write the <a:pPr> element for formula titles.
    #
    def write_a_p_pr_formula(font) # :nodoc:
      @writer.tag_elements('a:pPr') { write_a_def_rpr(font) }
    end

    #
    # Write the <a:defRPr> element.
    #
    def write_a_def_rpr(font = nil) # :nodoc:
      write_def_rpr_r_pr_common(
                                font,
                                get_font_style_attributes(font),
                                'a:defRPr')
    end

    #
    # Write the <a:endParaRPr> element.
    #
    def write_a_end_para_rpr # :nodoc:
      @writer.empty_tag('a:endParaRPr', ['lang', 'en-US'])
    end

    #
    # Write the <a:r> element.
    #
    def write_a_r(title, font) # :nodoc:
      @writer.tag_elements('a:r') do
        # Write the a:rPr element.
        write_a_r_pr(font)
        # Write the a:t element.
        write_a_t(title)
      end
    end

    #
    # Write the <a:rPr> element.
    #
    def write_a_r_pr(font) # :nodoc:
      write_def_rpr_r_pr_common(
                                font,
                                get_font_style_attributes(font).unshift('en-US').unshift('lang'),
                                'a:rPr'
                                )
    end

    def write_def_rpr_r_pr_common(font, style_attributes, tag)  # :nodoc:
      latin_attributes = get_font_latin_attributes(font)
      has_color = ptrue?(font) && ptrue?(font[:_color])

      if !latin_attributes.empty? || has_color
        @writer.tag_elements(tag, style_attributes) do
          if has_color
            write_a_solid_fill(:color => font[:_color])
          end
          if !latin_attributes.empty?
            write_a_latin(latin_attributes)
          end
        end
      else
        @writer.empty_tag(tag, style_attributes)
      end
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
    def write_tx_pr(horiz, font) # :nodoc:
      @writer.tag_elements('c:txPr') do
        # Write the a:bodyPr element.
        write_a_body_pr(horiz)
        # Write the a:lstStyle element.
        write_a_lst_style
        # Write the a:p element.
        write_a_p_formula(font)
      end
    end

    #
    # Write the <c:marker> element.
    #
    def write_marker(marker = nil) # :nodoc:
      marker ||= @default_marker

      return unless ptrue?(marker)
      return if ptrue?(marker[:automatic])

      @writer.tag_elements('c:marker') do
        # Write the c:symbol element.
        write_symbol(marker[:type])
        # Write the c:size element.
        size = marker[:size]
        write_marker_size(size) if ptrue?(size)
        # Write the c:spPr element.
        write_sp_pr(marker)
      end
    end

    #
    # Write the <c:marker> element without a sub-element.
    #
    def write_marker_value # :nodoc:
      return unless @default_marker

      @writer.empty_tag('c:marker', ['val', 1])
    end

    #
    # Write the <c:size> element.
    #
    def write_marker_size(val) # :nodoc:
      @writer.empty_tag('c:size', ['val', val])
    end

    #
    # Write the <c:symbol> element.
    #
    def write_symbol(val) # :nodoc:
      @writer.empty_tag('c:symbol', ['val', val])
    end

    #
    # Write the <c:spPr> element.
    #
    def write_sp_pr(series) # :nodoc:
      return if (!series.has_key?(:_line) || !ptrue?(series[:_line][:_defined])) &&
                (!series.has_key?(:_fill) || !ptrue?(series[:_fill][:_defined]))

      @writer.tag_elements('c:spPr') do
        # Write the fill elements for solid charts such as pie and bar.
        if series[:_fill] && series[:_fill][:_defined] != 0
          if ptrue?(series[:_fill][:none])
            # Write the a:noFill element.
            write_a_no_fill
          else
            # Write the a:solidFill element.
            write_a_solid_fill(series[:_fill])
          end
        end
        # Write the a:ln element.
        write_a_ln(series[:_line]) if series[:_line] && ptrue?(series[:_line][:_defined])
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
        if ptrue?(line[:none])
          # Write the a:noFill element.
          write_a_no_fill
        elsif ptrue?(line[:color])
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
      @writer.empty_tag('a:srgbClr', ['val', val])
    end

    #
    # Write the <a:prstDash> element.
    #
    def write_a_prst_dash(val) # :nodoc:
      @writer.empty_tag('a:prstDash', ['val', val])
    end

    #
    # Write the <c:trendline> element.
    #
    def write_trendline(trendline) # :nodoc:
      return unless trendline

      @writer.tag_elements('c:trendline') do
        # Write the c:name element.
        write_name(trendline[:name])
        # Write the c:spPr element.
        write_sp_pr(trendline)
        # Write the c:trendlineType element.
        write_trendline_type(trendline[:type])
        # Write the c:order element for polynomial trendlines.
        write_trendline_order(trendline[:order]) if trendline[:type] == 'poly'
        # Write the c:period element for moving average trendlines.
        write_period(trendline[:period]) if trendline[:type] == 'movingAvg'
        # Write the c:forward element.
        write_forward(trendline[:forward])
        # Write the c:backward element.
        write_backward(trendline[:backward])
      end
    end

    #
    # Write the <c:trendlineType> element.
    #
    def write_trendline_type(val) # :nodoc:
      @writer.empty_tag('c:trendlineType', ['val', val])
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
      @writer.empty_tag('c:order', ['val', val])
    end

    #
    # Write the <c:period> element.
    #
    def write_period(val = 2) # :nodoc:
      @writer.empty_tag('c:period', ['val', val])
    end

    #
    # Write the <c:forward> element.
    #
    def write_forward(val) # :nodoc:
      return unless val

      @writer.empty_tag('c:forward', ['val', val])
    end

    #
    # Write the <c:backward> element.
    #
    def write_backward(val) # :nodoc:
      return unless val

      @writer.empty_tag('c:backward', ['val', val])
    end

    #
    # Write the <c:hiLowLines> element.
    #
    def write_hi_low_lines # :nodoc:
      write_lines_base('c:hiLowLines', @hi_low_lines)
    end

    #
    # Write the <c:dropLines> elent.
    #
    def write_drop_lines
      write_lines_base('c:dropLines', @drop_lines)
    end

    #
    # used from write_drop_lines and write_hi_low_lines
    #
    def write_lines_base(tag, lines)
      return unless lines

      if lines[:_line] && ptrue?(lines[:_line][:_defined])
        @writer.tag_elements(tag) { write_sp_pr(lines) }
      else
        @writer.empty_tag(tag)
      end
    end

    #
    # Write the <c:overlap> element.
    #
    def write_overlap(val = nil) # :nodoc:
      return unless val

      @writer.empty_tag('c:overlap', ['val', val])
    end

    #
    # Write the <c:numCache> element.
    #
    def write_num_cache(data) # :nodoc:
      @writer.tag_elements('c:numCache') do

        # Write the c:formatCode element.
        write_format_code('General')

        # Write the c:ptCount element.
        write_pt_count(data.size)

        (0..data.size - 1).each do |i|
          token = data[i]

          # Write non-numeric data as 0.
          if token &&
              !(token.to_s =~ /^([+-]?)(?=\d|\.\d)\d*(\.\d*)?([Ee]([+-]?\d+))?$/)
            token = 0
          end

          # Write the c:pt element.
          write_pt(i, token)
        end
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
      @writer.empty_tag('c:ptCount', ['val', val])
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
    # Write the <c:dPt> elements.
    #
    def write_d_pt(points = nil)
      return unless ptrue?(points)

      index = -1
      points.each do |point|
        index += 1
        next unless ptrue?(point)

        write_d_pt_point(index, point)
      end
    end

    #
    # Write an individual <c:dPt> element.
    #
    def write_d_pt_point(index, point)
      @writer.tag_elements('c:dPt') do
        # Write the c:idx element.
        write_idx(index)
        # Write the c:spPr element.
        write_sp_pr(point)
      end
    end
    #
    # Write the <c:dLbls> element.
    #
    def write_d_lbls(labels) # :nodoc:
      return unless labels

      @writer.tag_elements('c:dLbls') do
        # Write the c:dLblPos element.
        write_d_lbl_pos(labels[:position]) if labels[:position]
        # Write the c:showVal element.
        write_show_val if labels[:value]
        # Write the c:showCatName element.
        write_show_cat_name if labels[:category]
        # Write the c:showSerName element.
        write_show_ser_name if labels[:series_name]
        # Write the c:showPercent element.
        write_show_percent if labels[:percentage]
        # Write the c:showLeaderLines element.
        write_show_leader_lines if labels[:leader_lines]
      end
    end

    #
    # Write the <c:showVal> element.
    #
    def write_show_val # :nodoc:
      @writer.empty_tag('c:showVal', ['val', 1])
    end

    #
    # Write the <c:showCatName> element.
    #
    def write_show_cat_name # :nodoc:
      @writer.empty_tag('c:showCatName', ['val', 1])
    end

    #
    # Write the <c:showSerName> element.
    #
    def write_show_ser_name # :nodoc:
      @writer.empty_tag('c:showSerName', ['val', 1])
    end

    #
    # Write the <c:showPercent> element.
    #
    def write_show_percent
      @writer.empty_tag('c:showPercent', ['val', 1])
    end

    #
    # Write the <c:showLeaderLines> element.
    #
    def write_show_leader_lines
      @writer.empty_tag('c:showLeaderLines', ['val', 1])
    end

    #
    # Write the <c:dLblPos> element.
    #
    def write_d_lbl_pos(val)
      @writer.empty_tag('c:dLblPos', ['val', val])
    end

    #
    # Write the <c:delete> element.
    #
    def write_delete(val) # :nodoc:
      @writer.empty_tag('c:delete', ['val', val])
    end

    #
    # Write the <c:invertIfNegative> element.
    #
    def write_c_invert_if_negative(invert = nil) # :nodoc:
      return unless ptrue?(invert)

      @writer.empty_tag('c:invertIfNegative', ['val', 1])
    end

    #
    # Write the axis font elements.
    #
    def write_axis_font(font) # :nodoc:
      return unless font

      @writer.tag_elements('c:txPr') do
        write_axis_body_pr(font[:_rotation])
        write_a_lst_style
        @writer.tag_elements('a:p') do
          write_a_p_pr_rich(font)
          write_a_end_para_rpr
        end
      end
    end

    #
    # Write the <a:latin> element.
    #
    def write_a_latin(args)  # :nodoc:
      @writer.empty_tag('a:latin', args)
    end

    #
    # Write the <c:dTable> element.
    #
    def write_d_table
      @table.write_d_table(@writer) if @table
    end

    #
    # Write the X and Y error bars.
    #
    def write_error_bars(error_bars)
      return unless ptrue?(error_bars)

      if error_bars[:_x_error_bars]
        write_err_bars('x', error_bars[:_x_error_bars])
      end
      if error_bars[:_y_error_bars]
        write_err_bars('y', error_bars[:_y_error_bars])
      end
    end

    #
    # Write the <c:errBars> element.
    #
    def write_err_bars(direction, error_bars)
      return unless ptrue?(error_bars)

      @writer.tag_elements('c:errBars') do
        # Write the c:errDir element.
        write_err_dir(direction)

        # Write the c:errBarType element.
        write_err_bar_type(error_bars[:_direction])

        # Write the c:errValType element.
        write_err_val_type(error_bars[:_type])

        unless ptrue?(error_bars[:_endcap])
          # Write the c:noEndCap element.
          write_no_end_cap
        end

        if error_bars[:_type] != 'stdErr'
          # Write the c:val element.
          write_error_val(error_bars[:_value])
        end

        # Write the c:spPr element.
        write_sp_pr(error_bars)
      end
    end

    #
    # Write the <c:errDir> element.
    #
    def write_err_dir(val)
      @writer.empty_tag('c:errDir', ['val', val])
    end

    #
    # Write the <c:errBarType> element.
    #
    def write_err_bar_type(val)
      @writer.empty_tag('c:errBarType', ['val', val])
    end

    #
    # Write the <c:errValType> element.
    #
    def write_err_val_type(val)
      @writer.empty_tag('c:errValType', ['val', val])
    end

    #
    # Write the <c:noEndCap> element.
    #
    def write_no_end_cap
      @writer.empty_tag('c:noEndCap', ['val', 1])
    end

    #
    # Write the <c:val> element.
    #
    def write_error_val(val)
      @writer.empty_tag('c:val', ['val', val])
    end

    #
    # Write the <c:upDownBars> element.
    #
    def write_up_down_bars
      return unless ptrue?(@up_down_bars)

      @writer.tag_elements('c:upDownBars') do
        # Write the c:gapWidth element.
        write_gap_width(150)

        # Write the c:upBars element.
        write_up_bars(@up_down_bars[:_up])

        # Write the c:downBars element.
        write_down_bars(@up_down_bars[:_down])
      end
    end

    #
    # Write the <c:gapWidth> element.
    #
    def write_gap_width(val = nil)
      return unless val

      @writer.empty_tag('c:gapWidth', ['val', val])
    end

    #
    # Write the <c:upBars> element.
    #
    def write_up_bars(format)
      write_bars_base('c:upBars', format)
    end

    #
    # Write the <c:upBars> element.
    #
    def write_down_bars(format)
      write_bars_base('c:downBars', format)
    end

    #
    # Write the <c:smooth> element.
    #
    def write_c_smooth(smooth)
      return unless ptrue?(smooth)

      attributes = ['val', 1]

      @writer.empty_tag('c:smooth', attributes)
    end

    def write_bars_base(tag, format)
      if ptrue?(format[:_line][:_defined]) || ptrue?(format[:_fill][:_defined])
        @writer.tag_elements(tag) { write_sp_pr(format) }
      else
        @writer.empty_tag(tag)
      end
    end

    def nil_or_max?(val)  # :nodoc:
      val.nil? || val == 'max'
    end
  end
end
