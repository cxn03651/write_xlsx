# -*- coding: utf-8 -*-
require 'write_xlsx/package/xml_writer_simple'
require 'write_xlsx/utility'
require 'write_xlsx/chart/axis'
require 'write_xlsx/chart/caption'
require 'write_xlsx/chart/series'

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
      @title             = Caption.new(self)
      @x_axis            = Axis.new(self)
      @y_axis            = Axis.new(self)
      @x2_axis           = Axis.new(self)
      @y2_axis           = Axis.new(self)
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
    def add_series(params)
      # Check that the required input has been specified.
      unless params.has_key?(:values)
        raise "Must specify ':values' in add_series"
      end

      if @requires_category != 0 && !params.has_key?(:categories)
        raise  "Must specify ':categories' in add_series for this chart type"
      end

      @series << Series.new(self, params)

      # Set the gap and overlap for Bar/Column charts.
      @series_gap     = params[:gap]     if params[:gap]
      @series_overlap = params[:overlap] if params[:overlap]
    end

    #
    # Set the properties of the x-axis.
    #
    def set_x_axis(params = {})
      @x_axis.merge_with_hash(params)
    end

    #
    # Set the properties of the Y-axis.
    #
    # The set_y_axis() method is used to set properties of the Y axis.
    # The properties that can be set are the same as for set_x_axis,
    #
    def set_y_axis(params = {})
      @y_axis.merge_with_hash(params)
    end

    #
    # Set the properties of the secondary X-axis.
    #
    def set_x2_axis(params = {})
      @x2_axis.merge_with_hash(params)
    end

    #
    # Set the properties of the secondary Y-axis.
    #
    def set_y2_axis(params = {})
      @y2_axis.merge_with_hash(params)
    end

    #
    # Set the properties of the chart title.
    #
    def set_title(params)
      @title.merge_with_hash(params)
    end

    #
    # Set the properties of the chart legend.
    #
    def set_legend(params)
      @legend_position = params[:position] || 'right'
      @legend_delete_series = params[:delete_series]
    end

    #
    # Set the properties of the chart plotarea.
    #
    def set_plotarea(params)
      # Convert the user defined properties to internal properties.
      @plotarea = area_properties(params)
    end

    #
    # Set the properties of the chart chartarea.
    #
    def set_chartarea(params)
      # Convert the user defined properties to internal properties.
      @chartarea = area_properties(params)
    end

    #
    # Set on of the 42 built-in Excel chart styles. The default style is 2.
    #
    def set_style(style_id = 2)
      style_id = 2 if style_id < 0 || style_id > 42
      @style_id = style_id
    end

    #
    # Set the option for displaying blank data in a chart. The default is 'gap'.
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
    def set_table(params = {})
      @table = Table.new(params)
    end

    #
    # Set properties for the chart up-down bars.
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
    def set_drop_lines(params = {})
      # Set the drop line properties.
      line = line_properties(params[:line])

      @drop_lines = { :_line => line }
    end

    #
    # Set properties for the chart high-low lines.
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

    def value_or_raise(hash, key, msg)
      raise "Unknown #{msg} '#{key}'" unless hash[key.to_sym]
      hash[key.to_sym]
    end

    #
    # Returns series which use the primary axes.
    #
    def get_primary_axes_series
      @series.reject {|s| s.y2_axis}
    end
    alias :primary_axes_series :get_primary_axes_series

    #
    # Returns series which use the secondary axes.
    #
    def get_secondary_axes_series
      @series.select {|s| s.y2_axis}
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
        if @title.formula
          write_title_formula(@title.formula, @title.data_id, nil, @title.name_font)
        elsif @title.name
          write_title_rich(@title.name, nil, @title.name_font)
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
        write_marker(series.marker)
        # Write the c:invertIfNegative element.
        write_c_invert_if_negative(series.invert_if_neg)
        # Write the c:dPt element.
        write_d_pt(series.points)
        # Write the c:dLbls element.
        write_d_lbls(series.labels)
        # Write the c:trendline element.
        write_trendline(series.trendline)
        # Write the c:errBars element.
        write_error_bars(series.error_bars)
        # Write the c:cat element.
        write_cat(series)
        # Write the c:val element.
        write_val(series)
        # Write the c:smooth element.
        write_c_smooth(series.smooth) if ptrue?(@smooth_allowed)
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
      if name = series.name_formula
        write_tx_formula(name, series.name_id)
      elsif name = series.name
        write_tx_value(name)
      end
    end

    #
    # Write the <c:cat> element.
    #
    def write_cat(series) # :nodoc:

      formula = series.categories
      data_id = series.cat_data_id

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
      write_val_base(series.values, series.val_data_id, 'c:val')
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
        if x_axis.formula
          write_title_formula(x_axis.formula, @x_axis.data_id, horiz, @x_axis.name_font)
        elsif x_axis.name
          write_title_rich(x_axis.name, horiz, x_axis.name_font)
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
        if y_axis.formula
          write_title_formula(y_axis.formula, y_axis.data_id, horiz, y_axis.name_font)
        elsif y_axis.name
          write_title_rich(y_axis.name, horiz, y_axis.name_font)
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
        if x_axis.formula
          write_title_formula(x_axis.formula, y_axis.data_id, horiz, x_axis.name_font)
        elsif x_axis.name
          write_title_rich(x_axis.name, horiz, x_axis.name_font)
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
        if x_axis.formula
          write_title_formula(x_axis.formula, x_axis.data_id, nil, x_axis.name_font)
        elsif x_axis.name
          write_title_rich(x_axis.name, nil, x_axis.name_font)
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
      line = series.respond_to?(:line) ? series.line : series[:_line]
      fill = series.respond_to?(:fill) ? series.fill : series[:_fill]

      return if (!line || !ptrue?(line[:_defined])) &&
        (!fill || !ptrue?(fill[:_defined]))

      @writer.tag_elements('c:spPr') do
        # Write the fill elements for solid charts such as pie and bar.
        if fill && fill[:_defined] != 0
          if ptrue?(fill[:none])
            # Write the a:noFill element.
            write_a_no_fill
          else
            # Write the a:solidFill element.
            write_a_solid_fill(fill)
          end
        end
        # Write the a:ln element.
        write_a_ln(line) if line && ptrue?(line[:_defined])
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
