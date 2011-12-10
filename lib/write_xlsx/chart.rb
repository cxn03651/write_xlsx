# -*- coding: utf-8 -*-
require 'write_xlsx/package/xml_writer_simple'
require 'write_xlsx/utility'

module Writexlsx
  class Chart
    include Utility

    attr_accessor :id
    attr_writer :index, :palette
    attr_reader :embedded, :formula_ids, :formula_data

    #
    # Factory method for returning chart objects based on their class type.
    #
    def self.factory(chart_subclass)
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

    def initialize(subtype)
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

      set_default_properties
    end

    def set_xml_writer(filename)
      @writer.set_xml_writer(filename)
    end

    #
    # Assemble and write the XML file.
    #
    def assemble_xml_file
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

      # Add the user supplied data to the internal structures.
      @series << {
        :_values       => values,
        :_categories   => categories,
        :_name         => name,
        :_name_formula => name_formula,
        :_name_id      => name_id,
        :_val_data_id  => val_id,
        :_cat_data_id  => cat_id,
        :_line         => line,
        :_fill         => fill,
        :_marker       => marker,
        :_trendline    => trendline,
        :_labels       => labels
      }
    end

    #
    # Set the properties of the X-axis.
    #
    def set_x_axis(params)
      name, name_formula = process_names(params[:name], params[:name_formula])

      data_id = get_data_id(name_formula, params[:data])

      @x_axis_name    = name
      @x_axis_formula = name_formula
      @x_axis_data_id = data_id
      @x_axis_reverse = params[:reverse]
    end

    ###############################################################################
    #
    # set_y_axis
    #
    # Set the properties of the Y-axis.
    #
    def set_y_axis(params)
      name, name_formula = process_names(params[:name], params[name_formula])
      data_id = get_data_id(name_formula, params[:data])

      @y_axis_name    = name
      @y_axis_formula = name_formula
      @y_axis_data_id = data_id
      @y_axis_reverse = params[:reverse]
    end

    #
    # Set the properties of the chart title.
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
    def set_legend(params)
      @legend_position = params[:position] || 'right'
      @legend_delete_series = params[:delete_series]
    end

    #
    # Set the properties of the chart plotarea.
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

    private

    #
    # Convert and aref of row col values to a range formula.
    #
    def aref_to_formula(data)
      # If it isn't an array ref it is probably a formula already.
      return data unless data.kind_of?(Array)
      xl_range_formula(*data)
    end

    #
    # Switch name and name_formula parameters if required.
    #
    def process_names(name = nil, name_formula = nil)
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
    def get_data_type(data)
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
    def get_data_id(formula, data)
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
    def get_color(color)
      # Convert a HTML style #RRGGBB color.
      if color and color =~ /^#[0-9a-fA-F]{6}$/
        color = color.sub(/^#/, '')
        return color.upperca
      end

      index = Format.get_color(color)

      # Set undefined colors to black.
      unless index
        index = 0x08;
        raise "Unknown color '#{color}' used in chart formatting."
      end

      get_palette_color(index)
    end

    #
    # Convert from an Excel internal colour index to a XML style #RRGGBB index
    # based on the default or user defined values in the Workbook palette.
    # Note: This version doesn't add an alpha channel.
    #
    def get_palette_color(index)
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
    def get_line_pattern(value)
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
    def get_line_weight(value)
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
    def get_line_properties(line)
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
    def get_fill_properties(fill)
      return { :_defined => 0 } unless fill

      fill[:_defined] = 1

      fill
    end

    #
    # Convert user defined marker properties to the structure required internally.
    #
    def get_marker_properties(marker)
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
    def get_trendline_properties(trendline)
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
    def get_labels_properties(labels)
      return nil unless labels

      return labels
    end

    #
    # Add a unique id for an axis.
    #
    def add_axis_id
      chart_id   = 1 + @id
      axis_count = 1 + @axis_ids.size

      axis_id = sprintf('5%03d%04d', chart_id, axis_count)

      @axis_ids << axis_id

      axis_id
    end

    #
    # Setup the default properties for a chart.
    #
    def set_default_properties
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
    def write_chart_space
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
    def write_lang
      val  = 'en-US'

      attributes = ['val', val]

      @writer.empty_tag('c:lang', attributes)
    end

    #
    # Write the <c:style> element.
    #
    def write_style
      style_id = @style_id

      # Don't write an element for the default style, 2.
      return if style_id == 2

      attributes = ['val', style_id]

      @writer.empty_tag('c:style', attributes)
    end

    #
    # Write the <c:chart> element.
    #
    def write_chart
      @writer.start_tag('c:chart')

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

      @writer.end_tag('c:chart')
    end

    #
    # Write the <c:plotArea> element.
    #
    def write_plot_area
      @writer.start_tag('c:plotArea')

      # Write the c:layout element.
      write_layout

      # Write the subclass chart type element.
      write_chart_type

      # Write the c:catAx element.
      write_cat_axis

      # Write the c:catAx element.
      write_val_axis

      @writer.end_tag('c:plotArea')
    end

    #
    # Write the <c:layout> element.
    #
    def write_layout
      @writer.empty_tag('c:layout')
    end

    #
    # Write the chart type element. This method should be overridden by the
    # subclasses.
    #
    def write_chart_type
    end

    #
    # Write the <c:grouping> element.
    #
    def write_grouping(val)
      attributes = ['val', val]
      @writer.empty_tag('c:grouping', attributes)
    end

    #
    # Write the series elements.
    #
    def write_series
      # Write each series with subelements.
      index = 0
      @series.each do |series|
        write_ser(index, series)
        index += 1
      end

      # Write the c:marker element.
      write_marker_value

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
    def write_ser(index, series)
      @writer.start_tag('c:ser')

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

      # Write the c:dLbls element.
      write_d_lbls(series[:labels])

      # Write the c:trendline element.
      write_trendline(series[:trendline])

      # Write the c:cat element.
      write_cat(series)

      # Write the c:val element.
      write_val(series)

      @writer.end_tag('c:ser')
    end

    #
    # Write the <c:idx> element.
    #
    def write_idx(val)
      attributes = ['val', val]

      @writer.empty_tag('c:idx', attributes)
    end

    #
    # Write the <c:order> element.
    #
    def write_order(val)
      attributes = ['val', val]

      @writer.empty_tag('c:order', attributes)
    end

    #
    # Write the series name.
    #
    def write_series_name(series)
      if name = series[:_name_formula]
        write_tx_formula(name, series[:_name_id])
      elsif name = series[:_name]
        write_tx_value(name)
      end
    end

    #
    # Write the <c:cat> element.
    #
    def write_cat(series)

      formula = series[:_categories]
      data_id = series[:_cat_data_id]

      data = @formula_data[data_id] if data_id

      # Ignore <c:cat> elements for charts without category values.
      return unless formula

      @has_category = 1

      @writer.start_tag('c:cat')

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

      @writer.end_tag('c:cat')
    end

    #
    # Write the <c:val> element.
    #
    def write_val(series)
      formula = series[:_values]
      data_id = series[:_val_data_id]
      data    = @formula_data[data_id]

      @writer.start_tag('c:val')

      # Check the type of cached data.
      type = get_data_type(data)

      if type == 'str'
        # Write the c:numRef element.
        write_str_ref(formula, data, type)
      else
        # Write the c:numRef element.
        write_num_ref(formula, data, type)
      end

      @writer.end_tag('c:val')
    end

    #
    # Write the <c:numRef> element.
    #
    def write_num_ref(formula, data, type)
      @writer.start_tag('c:numRef')

      # Write the c:f element.
      write_series_formula(formula)

      if type == 'num'
        # Write the c:numCache element.
        write_num_cache(data)
      elsif type == 'str'
        # Write the c:strCache element.
        write_str_cache(data)
      end

      @writer.end_tag('c:numRef')
    end

    #
    # Write the <c:strRef> element.
    #
    def write_str_ref(formula, data, type)
      @writer.start_tag('c:strRef')

      # Write the c:f element.
      write_series_formula(formula)

      if type == 'num'
        # Write the c:numCache element.
        write_num_cache(data)
      elsif type == 'str'
        # Write the c:strCache element.
        write_str_cache(data)
      end

      @writer.end_tag('c:strRef')
    end

    #
    # Write the <c:f> element.
    #
    def write_series_formula(formula)
      # Strip the leading '=' from the formula.
      formula = formula.sub(/^=/, '')

      @writer.data_element('c:f', formula)
    end

    #
    # Write the <c:axId> element.
    #
    def write_axis_id(val)
      attributes = ['val', val]

      @writer.empty_tag('c:axId', attributes)
    end

    #
    # Write the <c:catAx> element.
    #
    def write_cat_axis(position = nil)
      position  ||= @cat_axis_position
      horiz     = @horiz_cat_axis
      x_reverse = @x_axis_reverse
      y_reverse = @y_axis_reverse

      @writer.start_tag('c:catAx')

      write_axis_id(@axis_ids[0])

      # Write the c:scaling element.
      write_scaling(x_reverse)

      # Write the c:axPos element.
      write_axis_pos(position, y_reverse)

      # Write the axis title elements.
      if title = @x_axis_formula
        write_title_formula(title, @x_axis_data_id, horiz)
      elsif title = @x_axis_name
        write_title_rich(title, horiz)
      end

      # Write the c:numFmt element.
      write_num_fmt

      # Write the c:tickLblPos element.
      write_tick_label_pos('nextTo')

      # Write the c:crossAx element.
      write_cross_axis(@axis_ids[1])

      # Write the c:crosses element.
      write_crosses('autoZero')

      # Write the c:auto element.
      write_auto(1)

      # Write the c:labelAlign element.
      write_label_align('ctr')

      # Write the c:labelOffset element.
      write_label_offset(100)

      @writer.end_tag('c:catAx')
    end

    #
    # Write the <c:valAx> element.
    #
    # TODO. Maybe should have a _write_cat_val_axis method as well for scatter.
    #
    def write_val_axis(position = nil, hide_major_gridlines = nil)
      position ||= @val_axis_position
      horiz      = @horiz_val_axis
      x_reverse  = @x_axis_reverse
      y_reverse  = @y_axis_reverse

      @writer.start_tag('c:valAx')

      write_axis_id(@axis_ids[1])

      # Write the c:scaling element.
      write_scaling(y_reverse)

      # Write the c:axPos element.
      write_axis_pos(position, x_reverse)

      # Write the c:majorGridlines element.
      write_major_gridlines unless hide_major_gridlines

      # Write the axis title elements.
      if title = @y_axis_formula
        write_title_formula(title, @y_axis_data_id, horiz)
      elsif title = @y_axis_name
        write_title_rich(title, horiz)
      end

      # Write the c:numberFormat element.
      write_number_format

      # Write the c:tickLblPos element.
      write_tick_label_pos('nextTo')

      # Write the c:crossAx element.
      write_cross_axis(@axis_ids[0])

      # Write the c:crosses element.
      write_crosses('autoZero')

      # Write the c:crossBetween element.
      write_cross_between

      @writer.end_tag('c:valAx')
    end

    #
    # Write the <c:valAx> element. This is for the second valAx in scatter plots.
    #
    #
    def write_cat_val_axis(position, hide_major_gridlines)
      position ||= @val_axis_position
      horiz                = @horiz_val_axis
      x_reverse            = @x_axis_reverse
      y_reverse            = @y_axis_reverse

      @writer.start_tag('c:valAx')

      write_axis_id(@axis_ids[0])

      # Write the c:scaling element.
      write_scaling(x_reverse)

      # Write the c:axPos element.
      write_axis_pos(position, y_reverse)

      # Write the c:majorGridlines element.
      write_major_gridlines unless hide_major_gridlines

      # Write the axis title elements.
      if title = @x_axis_formula
        write_title_formula(title, @y_axis_data_id, horiz)
      elsif title = @x_axis_name
        write_title_rich(title, horiz)
      end

      # Write the c:numberFormat element.
      write_number_format

      # Write the c:tickLblPos element.
      write_tick_label_pos('nextTo')

      # Write the c:crossAx element.
      write_cross_axis(@axis_ids[1])

      # Write the c:crosses element.
      write_crosses('autoZero')

      # Write the c:crossBetween element.
      write_cross_between

      @writer.end_tag('c:valAx')
    end

    #
    # Write the <c:dateAx> element.
    #
    def write_date_axis(position = nil)
      position ||= @cat_axis_position
      x_reverse  = @x_axis_reverse
      y_reverse  = @y_axis_reverse

      @writer.start_tag('c:dateAx')

      write_axis_id(@axis_ids[0])

      # Write the c:scaling element.
      write_scaling(x_reverse)

      # Write the c:axPos element.
      write_axis_pos(position, y_reverse)

      # Write the axis title elements.
      if title = @x_axis_formula
        write_title_formula(title, @x_axis_data_id)
      elsif title = @x_axis_name
        write_title_rich(title)
      end

      # Write the c:numFmt element.
      write_num_fmt('dd/mm/yyyy')

      # Write the c:tickLblPos element.
      write_tick_label_pos('nextTo')

      # Write the c:crossAx element.
      write_cross_axis(@axis_ids[1])

      # Write the c:crosses element.
      write_crosses('autoZero')

      # Write the c:auto element.
      write_auto(1)

      # Write the c:labelOffset element.
      write_label_offset(100)

      @writer.end_tag('c:dateAx')
    end

    #
    # Write the <c:scaling> element.
    #
    def write_scaling(reverse)
      @writer.start_tag('c:scaling')

      # Write the c:orientation element.
      write_orientation(reverse)

      @writer.end_tag('c:scaling')
    end

    #
    # Write the <c:orientation> element.
    #
    def write_orientation(reverse = nil)
      val     = reverse ? 'maxMin' : 'minMax'

      attributes = ['val', val]

      @writer.empty_tag('c:orientation', attributes)
    end

    #
    # Write the <c:axPos> element.
    #
    def write_axis_pos(val, reverse = false)
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
    def write_num_fmt(format_code = nil)
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
    def write_tick_label_pos(val)
      attributes = ['val', val]

      @writer.empty_tag('c:tickLblPos', attributes)
    end

    #
    # Write the <c:crossAx> element.
    #
    def write_cross_axis(val)
      attributes = ['val', val]

      @writer.empty_tag('c:crossAx', attributes)
    end

    #
    # Write the <c:crosses> element.
    #
    def write_crosses(val)
      attributes = ['val', val]

      @writer.empty_tag('c:crosses', attributes)
    end

    #
    # Write the <c:auto> element.
    #
    def write_auto(val)
      attributes = ['val', val]

      @writer.empty_tag('c:auto', attributes)
    end

    #
    # Write the <c:labelAlign> element.
    #
    def write_label_align(val)
      attributes = ['val', val]

      @writer.empty_tag('c:lblAlgn', attributes)
    end

    #
    # Write the <c:labelOffset> element.
    #
    def write_label_offset(val)
      attributes = ['val', val]

      @writer.empty_tag('c:lblOffset', attributes)
    end

    #
    # Write the <c:majorGridlines> element.
    #
    def write_major_gridlines
      @writer.empty_tag('c:majorGridlines')
    end

    #
    # Write the <c:numberFormat> element.
    #
    # TODO. Merge/replace with _write_num_fmt.
    #
    def write_number_format
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
    def write_cross_between
      val  = @cross_between || 'between'

      attributes = ['val', val]

      @writer.empty_tag('c:crossBetween', attributes)
    end

    #
    # Write the <c:legend> element.
    #
    def write_legend
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

      @writer.start_tag('c:legend')

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

      @writer.end_tag('c:legend')
    end

    #
    # Write the <c:legendPos> element.
    #
    def write_legend_pos(val)
      attributes = ['val', val]

      @writer.empty_tag('c:legendPos', attributes)
    end

    #
    # Write the <c:legendEntry> element.
    #
    def write_legend_entry(index)
      @writer.start_tag('c:legendEntry')

      # Write the c:idx element.
      write_idx(index)

      # Write the c:delete element.
      write_delete(1)

      @writer.end_tag('c:legendEntry')
    end

    #
    # Write the <c:overlay> element.
    #
    def write_overlay
      val  = 1

      attributes = ['val', val]

      @writer.empty_tag('c:overlay', attributes)
    end

    #
    # Write the <c:plotVisOnly> element.
    #
    def write_plot_vis_only
      val  = 1

      attributes = ['val', val]

      @writer.empty_tag('c:plotVisOnly', attributes)
    end

    #
    # Write the <c:printSettings> element.
    #
    def write_print_settings
      @writer.start_tag('c:printSettings')

      # Write the c:headerFooter element.
      write_header_footer

      # Write the c:pageMargins element.
      write_page_margins

      # Write the c:pageSetup element.
      write_page_setup

      @writer.end_tag('c:printSettings')
    end

    #
    # Write the <c:headerFooter> element.
    #
    def write_header_footer
      @writer.empty_tag('c:headerFooter')
    end

    #
    # Write the <c:pageMargins> element.
    #
    def write_page_margins
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
    def write_page_setup
      @writer.empty_tag('c:pageSetup')
    end

    #
    # Write the <c:title> element for a rich string.
    #
    def write_title_rich(title, horiz = nil)
      @writer.start_tag('c:title')

      # Write the c:tx element.
      write_tx_rich(title, horiz)

      # Write the c:layout element.
      write_layout

      @writer.end_tag('c:title')
    end

    #
    # Write the <c:title> element for a rich string.
    #
    def write_title_formula(title, data_id, horiz)
      @writer.start_tag('c:title')

      # Write the c:tx element.
      write_tx_formula(title, data_id)

      # Write the c:layout element.
      write_layout

      # Write the c:txPr element.
      write_tx_pr(horiz)

      @writer.end_tag('c:title')
    end

    #
    # Write the <c:tx> element.
    #
    def write_tx_rich(title, horiz)
      @writer.start_tag('c:tx')

      # Write the c:rich element.
      write_rich(title, horiz)

      @writer.end_tag('c:tx')
    end

    #
    # Write the <c:tx> element with a simple value such as for series names.
    #
    def write_tx_value(title)
      @writer.start_tag('c:tx')

      # Write the c:v element.
      write_v(title)

      @writer.end_tag('c:tx')
    end

    #
    # Write the <c:tx> element.
    #
    def write_tx_formula(title, data_id)
      data = @formula_data[data_id] if data_id

      @writer.start_tag('c:tx')

      # Write the c:strRef element.
      write_str_ref(title, data, 'str')

      @writer.end_tag('c:tx')
    end

    #
    # Write the <c:rich> element.
    #
    def write_rich(title, horiz)
      @writer.start_tag('c:rich')

      # Write the a:bodyPr element.
      write_a_body_pr(horiz)

      # Write the a:lstStyle element.
      write_a_lst_style

      # Write the a:p element.
      write_a_p_rich(title)

      @writer.end_tag('c:rich')
    end

    #
    # Write the <a:bodyPr> element.
    #
    def write_a_body_pr(horiz)
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
    def write_a_lst_style
      @writer.empty_tag('a:lstStyle')
    end

    #
    # Write the <a:p> element for rich string titles.
    #
    def write_a_p_rich(title)
      @writer.start_tag('a:p')

      # Write the a:pPr element.
      write_a_p_pr_rich

      # Write the a:r element.
      write_a_r(title)

      @writer.end_tag('a:p')
    end

    #
    # Write the <a:p> element for formula titles.
    #
    def write_a_p_formula(title)
      @writer.start_tag('a:p')

      # Write the a:pPr element.
      write_a_p_pr_formula

      # Write the a:endParaRPr element.
      write_a_end_para_rpr

      @writer.end_tag('a:p')
    end

    #
    # Write the <a:pPr> element for rich string titles.
    #
    def write_a_p_pr_rich
      @writer.start_tag('a:pPr')

      # Write the a:defRPr element.
      write_a_def_rpr

      @writer.end_tag('a:pPr')
    end

    #
    # Write the <a:pPr> element for formula titles.
    #
    def write_a_p_pr_formula
      @writer.start_tag('a:pPr')

      # Write the a:defRPr element.
      write_a_def_rpr

      @writer.end_tag('a:pPr')
    end

    #
    # Write the <a:defRPr> element.
    #
    def write_a_def_rpr
      @writer.empty_tag('a:defRPr')
    end

    #
    # Write the <a:endParaRPr> element.
    #
    def write_a_end_para_rpr
      lang = 'en-US'

      attributes = ['lang', lang]

      @writer.empty_tag('a:endParaRPr', attributes)
    end

    #
    # Write the <a:r> element.
    #
    def write_a_r(title)
      @writer.start_tag('a:r')

      # Write the a:rPr element.
      write_a_r_pr

      # Write the a:t element.
      write_a_t(title)

      @writer.end_tag('a:r')
    end

    #
    # Write the <a:rPr> element.
    #
    def write_a_r_pr
      lang = 'en-US'

      attributes = ['lang', lang]

      @writer.empty_tag('a:rPr', attributes)
    end

    #
    # Write the <a:t> element.
    #
    def write_a_t(title)
      @writer.data_element('a:t', title)
    end

    #
    # Write the <c:txPr> element.
    #
    def write_tx_pr(horiz)
      @writer.start_tag('c:txPr')

      # Write the a:bodyPr element.
      write_a_body_pr(horiz)

      # Write the a:lstStyle element.
      write_a_lst_style

      # Write the a:p element.
      write_a_p_formula

      @writer.end_tag('c:txPr')
    end

    #
    # Write the <c:marker> element.
    #
    def write_marker(marker = nil)
      marker ||= @default_marker

      return if marker.nil? || marker == 0
      return if marker[:automatic] && marker[:automatic] != 0

      @writer.start_tag('c:marker')

      # Write the c:symbol element.
      write_symbol(marker[:type])

      # Write the c:size element.
      size = marker[:size]
      write_marker_size(size) if !size.nil? && size != 0

      # Write the c:spPr element.
      write_sp_pr(marker)

      @writer.end_tag('c:marker')
    end

    #
    # Write the <c:marker> element without a sub-element.
    #
    def write_marker_value
      style = @default_marker

      return unless style

      attributes = ['val', 1]

      @writer.empty_tag('c:marker', attributes)
    end

    #
    # Write the <c:size> element.
    #
    def write_marker_size(val)
      attributes = ['val', val]

      @writer.empty_tag('c:size', attributes)
    end

    #
    # Write the <c:symbol> element.
    #
    def write_symbol(val)
      attributes = ['val', val]

      @writer.empty_tag('c:symbol', attributes)
    end

    #
    # Write the <c:spPr> element.
    #
    def write_sp_pr(series)
      return if (!series.has_key?(:_line) || series[:_line][:_defined].nil? || series[:_line][:_defined]== 0) &&
                (!series.has_key?(:_fill) || series[:_fill][:_defined].nil? || series[:_fill][:_defined]== 0)

      @writer.start_tag('c:spPr')

      # Write the a:solidFill element for solid charts such as pie and bar.
      write_a_solid_fill(series[:_fill]) if series[:_fill] && series[:_fill][:_defined] != 0

      # Write the a:ln element.
      write_a_ln(series[:_line]) if series[:_line] && series[:_line][:_defined]

      @writer.end_tag('c:spPr')
    end

    #
    # Write the <a:ln> element.
    #
    def write_a_ln(line)
      attributes = []

      # Add the line width as an attribute.
      if width = line[:width]
        # Round width to nearest 0.25, like Excel.
        width = ((width + 0.125) * 4).to_i / 4.0

        # Convert to internal units.
        width = (0.5 + (12700 * width)).to_i

        attributes = ['w', width]
      end

      @writer.start_tag('a:ln', attributes)

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

      @writer.end_tag('a:ln')
    end

    #
    # Write the <a:noFill> element.
    #
    def write_a_no_fill
      @writer.empty_tag('a:noFill')
    end

    #
    # Write the <a:solidFill> element.
    #
    def write_a_solid_fill(line)
      @writer.start_tag('a:solidFill')

      if line[:color]
        color = get_color(line[:color])

        # Write the a:srgbClr element.
        write_a_srgb_clr(color)
      end

      @writer.end_tag('a:solidFill')
    end

    #
    # Write the <a:srgbClr> element.
    #
    def write_a_srgb_clr(val)
      attributes = ['val', val]

      @writer.empty_tag('a:srgbClr', attributes)
    end

    #
    # Write the <a:prstDash> element.
    #
    def write_a_prst_dash(val)
      attributes = ['val', val]

      @writer.empty_tag('a:prstDash', attributes)
    end

    #
    # Write the <c:trendline> element.
    #
    def write_trendline(trendline)
      return unless trendline

      @writer.start_tag('c:trendline')

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

      @writer.end_tag('c:trendline')
    end

    #
    # Write the <c:trendlineType> element.
    #
    def write_trendline_type(val)
      attributes = ['val', val]

      @writer.empty_tag('c:trendlineType', attributes)
    end

    #
    # Write the <c:name> element.
    #
    def write_name(data)
      return unless data

      @writer.data_element('c:name', data)
    end

    #
    # Write the <c:order> element.
    #
    def write_trendline_order(val)
      attributes = ['val', val]

      @writer.empty_tag('c:order', attributes)
    end

    #
    # Write the <c:period> element.
    #
    def write_period(val)
      val  ||= 2

      attributes = ['val', val]

      @writer.empty_tag('c:period', attributes)
    end

    #
    # Write the <c:forward> element.
    #
    def write_forward(val)
      return unless val

      attributes = ['val', val]

      @writer.empty_tag('c:forward', attributes)
    end

    #
    # Write the <c:backward> element.
    #
    def write_backward(val)
      return unless val

      attributes = ['val', val]

      @writer.empty_tag('c:backward', attributes)
    end

    #
    # Write the <c:hiLowLines> element.
    #
    def write_hi_low_lines
      @writer.empty_tag('c:hiLowLines')
    end

    #
    # Write the <c:overlap> element.
    #
    def write_overlap
      val  = 100

      attributes = ['val', val]

      @writer.empty_tag('c:overlap', attributes)
    end

    #
    # Write the <c:numCache> element.
    #
    def write_num_cache(data)
      count = data.size

      @writer.start_tag('c:numCache')

      # Write the c:formatCode element.
      write_format_code('General')

      # Write the c:ptCount element.
      write_pt_count(count)

      (0 .. count - 1).each do |i|

        # Write the c:pt element.
        write_pt(i, data[i])
      end

      @writer.end_tag('c:numCache')
    end

    #
    # Write the <c:strCache> element.
    #
    def write_str_cache(data)
      count = data.size

      @writer.start_tag('c:strCache')

      # Write the c:ptCount element.
      write_pt_count(count)

      (0 .. count - 1).each do |i|

        # Write the c:pt element.
        write_pt(i, data[i])
      end

      @writer.end_tag('c:strCache')
    end

    #
    # Write the <c:formatCode> element.
    #
    def write_format_code(data)
      @writer.data_element('c:formatCode', data)
    end

    #
    # Write the <c:ptCount> element.
    #
    def write_pt_count(val)
      attributes = ['val', val]

      @writer.empty_tag('c:ptCount', attributes)
    end

    #
    # Write the <c:pt> element.
    #
    def write_pt(idx, value)
      return unless value

      attributes = ['idx', idx]

      @writer.start_tag('c:pt', attributes)

      # Write the c:v element.
      write_v(value)

      @writer.end_tag('c:pt')
    end

    #
    # Write the <c:v> element.
    #
    def write_v(data)
      @writer.data_element('c:v', data)
    end

    #
    # Write the <c:protection> element.
    #
    def write_protection
      return if @protection == 0

      @writer.empty_tag('c:protection')
    end

    #
    # Write the <c:dLbls> element.
    #
    def write_d_lbls(labels)
      return unless labels

      @writer.start_tag('c:dLbls')

      # Write the c:showVal element.
      write_show_val if labels[value]

      # Write the c:showCatName element.
      write_show_cat_name if labels[category]

      # Write the c:showSerName element.
      write_show_ser_name if labels[series_name]

      @writer.end_tag('c:dLbls')
    end

    #
    # Write the <c:showVal> element.
    #
    def write_show_val
      val  = 1

      attributes = ['val', val]

      @writer.empty_tag('c:showVal', attributes)
    end

    #
    # Write the <c:showCatName> element.
    #
    def write_show_cat_name
      val  = 1

      attributes = ['val', val]

      @writer.empty_tag('c:showCatName', attributes)
    end

    #
    # Write the <c:showSerName> element.
    #
    def write_show_ser_name
      val  = 1

      attributes = ['val', val]

      @writer.empty_tag('c:showSerName', attributes)
    end

    #
    # Write the <c:delete> element.
    #
    def write_delete(val)
      attributes = ['val', val]

      @writer.empty_tag('c:delete', attributes)
    end
  end
end
