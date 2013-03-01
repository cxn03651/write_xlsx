# -*- coding: utf-8 -*-

module Writexlsx
  ###############################################################################
  #
  # Sparkline - A class for writing Excel shapes.
  #
  # Used in conjunction with Excel::Writer::XLSX.
  #
  # Copyright 2000-2012, John McNamara, jmcnamara@cpan.org
  # Converted to ruby by Hideo NAKAMURA, cxn03651@msj.biglobe.ne.jp
  #
  class Sparkline
    include Writexlsx::Utility

    attr_accessor :locations, :ranges
    attr_accessor :date_axis, :series_color, :negative_color, :markers_color
    attr_accessor :first_color, :last_color, :high_color, :low_color

    def initialize(ws, param, sheetname)
      @color = {}

      # Check for valid input parameters.
      param.each_key do |k|
        unless valid_sparkline_parameter[k]
          raise "Unknown parameter '#{k}' in add_sparkline()"
        end
      end
      [:location, :range].each do |required_key|
        unless param[required_key]
          raise "Parameter '#{required_key}' is required in add_sparkline()"
        end
      end

      # Handle the sparkline type.
      type = param[:type] || 'line'
      unless ['line', 'column', 'win_loss'].include?(type)
        raise "Parameter ':type' must be 'line', 'column' or 'win_loss' in add_sparkline()"
      end
      type = 'stacked' if type == 'win_loss'
      @type = type

      # We handle single location/range values or array refs of values.
      @locations = [param[:location]].flatten
      @ranges    = [param[:range]].flatten

      if @ranges.size != @locations.size
        raise "Must have the same number of location and range parameters in add_sparkline()"
      end

      # Store the count.
      @count = @locations.size

      # Cleanup the input ranges.
      @ranges.collect! do |range|
        # Remove the absolute reference $ symbols.
        range = range.gsub(/\$/, '')
        # Convert a simple range into a full Sheet1!A1:D1 range.
        range = "#{sheetname}!#{range}" unless range =~ /!/
        range
      end
      # Cleanup the input locations.
      @locations.collect! { |location| location.gsub(/\$/, '') }

      # Map options.
      @high      = param[:high_point]
      @low       = param[:low_point]
      @negative  = param[:negative_points]
      @first     = param[:first_point]
      @last      = param[:last_point]
      @markers   = param[:markers]
      @min       = param[:min]
      @max       = param[:max]
      @axis      = param[:axis]
      @reverse   = param[:reverse]
      @hidden    = param[:show_hidden]
      @weight    = param[:weight]

      # Map empty cells options.
      empty = param[:empty_cells] || ''
      @empty = case empty
      when 'zero'
        0
      when 'connect'
        'span'
      else
        'gap'
      end

      # Map the date axis range.
      date_range = param[:date_axis]
      if ptrue?(date_range) && !(date_range =~ /!/)
        date_range = "#{sheetname}!#{date_range}"
      end
      @date_axis = date_range

      # Set the sparkline styles.
      style_id = param[:style] || 0
      style = spark_styles[style_id]

      @series_color   = style[:series]
      @negative_color = style[:negative]
      @markers_color  = style[:markers]
      @first_color    = style[:first]
      @last_color     = style[:last]
      @high_color     = style[:high]
      @low_color      = style[:low]

      # Override the style colours with user defined colors.
      [:series_color, :negative_color, :markers_color, :first_color, :last_color, :high_color, :low_color].each do |user_color|
        set_spark_color(user_color, ptrue?(param[user_color]) ? ws.get_palette_color(param[user_color]) : nil)
      end
    end

    def group_attributes
      @cust_max = cust_max_min(@max) if @max
      @cust_min = cust_max_min(@min) if @min

      @cust_max = cust_max_min(@max) if @max
      @cust_min = cust_max_min(@min) if @min

      a = []
      a << 'manualMax' << @max if @max && @max != 'group'
      a << 'manualMin' << @min if @min && @min != 'group'

      # Ignore the default type attribute (line).
      a << 'type'          << @type        if @type != 'line'

      a << 'lineWeight'    << @weight      if @weight
      a << 'dateAxis'      << 1            if @date_axis
      a << 'displayEmptyCellsAs' << @empty if ptrue?(@empty)

      a << 'markers'       << 1          if @markers
      a << 'high'          << 1          if @high
      a << 'low'           << 1          if @low
      a << 'first'         << 1          if @first
      a << 'last'          << 1          if @last
      a << 'negative'      << 1          if @negative
      a << 'displayXAxis'  << 1          if @axis
      a << 'displayHidden' << 1          if @hidden
      a << 'minAxisType'   << @cust_min  if @cust_min
      a << 'maxAxisType'   << @cust_max  if @cust_max
      a << 'rightToLeft'   << 1          if @reverse
      a
    end

    private

    def set_spark_color(user_color, palette_color)
      return unless palette_color

      instance_variable_set("@#{user_color}", { :_rgb => palette_color })
    end

    def cust_max_min(max_min)  # :nodoc:
      max_min == 'group' ? 'group' : 'custom'
    end

    def valid_sparkline_parameter  # :nodoc:
      {
        :location        => 1,
        :range           => 1,
        :type            => 1,
        :high_point      => 1,
        :low_point       => 1,
        :negative_points => 1,
        :first_point     => 1,
        :last_point      => 1,
        :markers         => 1,
        :style           => 1,
        :series_color    => 1,
        :negative_color  => 1,
        :markers_color   => 1,
        :first_color     => 1,
        :last_color      => 1,
        :high_color      => 1,
        :low_color       => 1,
        :max             => 1,
        :min             => 1,
        :axis            => 1,
        :reverse         => 1,
        :empty_cells     => 1,
        :show_hidden     => 1,
        :date_axis       => 1,
        :weight          => 1
      }
    end

    def spark_styles  # :nodoc:
      [
       {   # 0
         :series   => { :_theme => "4", :_tint => "-0.499984740745262" },
         :negative => { :_theme => "5" },
         :markers  => { :_theme => "4", :_tint => "-0.499984740745262" },
         :first    => { :_theme => "4", :_tint => "0.39997558519241921" },
         :last     => { :_theme => "4", :_tint => "0.39997558519241921" },
         :high     => { :_theme => "4" },
         :low      => { :_theme => "4" }
       },
       {   # 1
         :series   => { :_theme => "4", :_tint => "-0.499984740745262" },
         :negative => { :_theme => "5" },
         :markers  => { :_theme => "4", :_tint => "-0.499984740745262" },
         :first    => { :_theme => "4", :_tint => "0.39997558519241921" },
         :last     => { :_theme => "4", :_tint => "0.39997558519241921" },
         :high     => { :_theme => "4" },
         :low      => { :_theme => "4" }
       },
       {   # 2
         :series   => { :_theme => "5", :_tint => "-0.499984740745262" },
         :negative => { :_theme => "6" },
         :markers  => { :_theme => "5", :_tint => "-0.499984740745262" },
         :first    => { :_theme => "5", :_tint => "0.39997558519241921" },
         :last     => { :_theme => "5", :_tint => "0.39997558519241921" },
         :high     => { :_theme => "5" },
         :low      => { :_theme => "5" }
       },
       {   # 3
         :series   => { :_theme => "6", :_tint => "-0.499984740745262" },
         :negative => { :_theme => "7" },
         :markers  => { :_theme => "6", :_tint => "-0.499984740745262" },
         :first    => { :_theme => "6", :_tint => "0.39997558519241921" },
         :last     => { :_theme => "6", :_tint => "0.39997558519241921" },
         :high     => { :_theme => "6" },
         :low      => { :_theme => "6" }
       },
       {   # 4
         :series   => { :_theme => "7", :_tint => "-0.499984740745262" },
         :negative => { :_theme => "8" },
         :markers  => { :_theme => "7", :_tint => "-0.499984740745262" },
         :first    => { :_theme => "7", :_tint => "0.39997558519241921" },
         :last     => { :_theme => "7", :_tint => "0.39997558519241921" },
         :high     => { :_theme => "7" },
         :low      => { :_theme => "7" }
       },
       {   # 5
         :series   => { :_theme => "8", :_tint => "-0.499984740745262" },
         :negative => { :_theme => "9" },
         :markers  => { :_theme => "8", :_tint => "-0.499984740745262" },
         :first    => { :_theme => "8", :_tint => "0.39997558519241921" },
         :last     => { :_theme => "8", :_tint => "0.39997558519241921" },
         :high     => { :_theme => "8" },
         :low      => { :_theme => "8" }
       },
       {   # 6
         :series   => { :_theme => "9", :_tint => "-0.499984740745262" },
         :negative => { :_theme => "4" },
         :markers  => { :_theme => "9", :_tint => "-0.499984740745262" },
         :first    => { :_theme => "9", :_tint => "0.39997558519241921" },
         :last     => { :_theme => "9", :_tint => "0.39997558519241921" },
         :high     => { :_theme => "9" },
         :low      => { :_theme => "9" }
       },
       {   # 7
         :series   => { :_theme => "4", :_tint => "-0.249977111117893" },
         :negative => { :_theme => "5" },
         :markers  => { :_theme => "5", :_tint => "-0.249977111117893" },
         :first    => { :_theme => "5", :_tint => "-0.249977111117893" },
         :last     => { :_theme => "5", :_tint => "-0.249977111117893" },
         :high     => { :_theme => "5", :_tint => "-0.249977111117893" },
         :low      => { :_theme => "5", :_tint => "-0.249977111117893" }
       },
       {   # 8
         :series   => { :_theme => "5", :_tint => "-0.249977111117893" },
         :negative => { :_theme => "6" },
         :markers  => { :_theme => "6", :_tint => "-0.249977111117893" },
         :first    => { :_theme => "6", :_tint => "-0.249977111117893" },
         :last     => { :_theme => "6", :_tint => "-0.249977111117893" },
         :high     => { :_theme => "6", :_tint => "-0.249977111117893" },
         :low      => { :_theme => "6", :_tint => "-0.249977111117893" }
       },
       {   # 9
         :series   => { :_theme => "6", :_tint => "-0.249977111117893" },
         :negative => { :_theme => "7" },
         :markers  => { :_theme => "7", :_tint => "-0.249977111117893" },
         :first    => { :_theme => "7", :_tint => "-0.249977111117893" },
         :last     => { :_theme => "7", :_tint => "-0.249977111117893" },
         :high     => { :_theme => "7", :_tint => "-0.249977111117893" },
         :low      => { :_theme => "7", :_tint => "-0.249977111117893" }
       },
       {   # 10
         :series   => { :_theme => "7", :_tint => "-0.249977111117893" },
         :negative => { :_theme => "8" },
         :markers  => { :_theme => "8", :_tint => "-0.249977111117893" },
         :first    => { :_theme => "8", :_tint => "-0.249977111117893" },
         :last     => { :_theme => "8", :_tint => "-0.249977111117893" },
         :high     => { :_theme => "8", :_tint => "-0.249977111117893" },
         :low      => { :_theme => "8", :_tint => "-0.249977111117893" }
       },
       {   # 11
         :series   => { :_theme => "8", :_tint => "-0.249977111117893" },
         :negative => { :_theme => "9" },
         :markers  => { :_theme => "9", :_tint => "-0.249977111117893" },
         :first    => { :_theme => "9", :_tint => "-0.249977111117893" },
         :last     => { :_theme => "9", :_tint => "-0.249977111117893" },
         :high     => { :_theme => "9", :_tint => "-0.249977111117893" },
         :low      => { :_theme => "9", :_tint => "-0.249977111117893" }
       },
       {   # 12
         :series   => { :_theme => "9", :_tint => "-0.249977111117893" },
         :negative => { :_theme => "4" },
         :markers  => { :_theme => "4", :_tint => "-0.249977111117893" },
         :first    => { :_theme => "4", :_tint => "-0.249977111117893" },
         :last     => { :_theme => "4", :_tint => "-0.249977111117893" },
         :high     => { :_theme => "4", :_tint => "-0.249977111117893" },
         :low      => { :_theme => "4", :_tint => "-0.249977111117893" }
       },
       {   # 13
         :series   => { :_theme => "4" },
         :negative => { :_theme => "5" },
         :markers  => { :_theme => "4", :_tint => "-0.249977111117893" },
         :first    => { :_theme => "4", :_tint => "-0.249977111117893" },
         :last     => { :_theme => "4", :_tint => "-0.249977111117893" },
         :high     => { :_theme => "4", :_tint => "-0.249977111117893" },
         :low      => { :_theme => "4", :_tint => "-0.249977111117893" }
       },
       {   # 14
         :series   => { :_theme => "5" },
         :negative => { :_theme => "6" },
         :markers  => { :_theme => "5", :_tint => "-0.249977111117893" },
         :first    => { :_theme => "5", :_tint => "-0.249977111117893" },
         :last     => { :_theme => "5", :_tint => "-0.249977111117893" },
         :high     => { :_theme => "5", :_tint => "-0.249977111117893" },
         :low      => { :_theme => "5", :_tint => "-0.249977111117893" }
       },
       {   # 15
         :series   => { :_theme => "6" },
         :negative => { :_theme => "7" },
         :markers  => { :_theme => "6", :_tint => "-0.249977111117893" },
         :first    => { :_theme => "6", :_tint => "-0.249977111117893" },
         :last     => { :_theme => "6", :_tint => "-0.249977111117893" },
         :high     => { :_theme => "6", :_tint => "-0.249977111117893" },
         :low      => { :_theme => "6", :_tint => "-0.249977111117893" }
       },
       {   # 16
         :series   => { :_theme => "7" },
         :negative => { :_theme => "8" },
         :markers  => { :_theme => "7", :_tint => "-0.249977111117893" },
         :first    => { :_theme => "7", :_tint => "-0.249977111117893" },
         :last     => { :_theme => "7", :_tint => "-0.249977111117893" },
         :high     => { :_theme => "7", :_tint => "-0.249977111117893" },
         :low      => { :_theme => "7", :_tint => "-0.249977111117893" }
       },
       {   # 17
         :series   => { :_theme => "8" },
         :negative => { :_theme => "9" },
         :markers  => { :_theme => "8", :_tint => "-0.249977111117893" },
         :first    => { :_theme => "8", :_tint => "-0.249977111117893" },
         :last     => { :_theme => "8", :_tint => "-0.249977111117893" },
         :high     => { :_theme => "8", :_tint => "-0.249977111117893" },
         :low      => { :_theme => "8", :_tint => "-0.249977111117893" }
       },
       {   # 18
         :series   => { :_theme => "9" },
         :negative => { :_theme => "4" },
         :markers  => { :_theme => "9", :_tint => "-0.249977111117893" },
         :first    => { :_theme => "9", :_tint => "-0.249977111117893" },
         :last     => { :_theme => "9", :_tint => "-0.249977111117893" },
         :high     => { :_theme => "9", :_tint => "-0.249977111117893" },
         :low      => { :_theme => "9", :_tint => "-0.249977111117893" }
       },
       {   # 19
         :series   => { :_theme => "4", :_tint => "0.39997558519241921" },
         :negative => { :_theme => "0", :_tint => "-0.499984740745262" },
         :markers  => { :_theme => "4", :_tint => "0.79998168889431442" },
         :first    => { :_theme => "4", :_tint => "-0.249977111117893" },
         :last     => { :_theme => "4", :_tint => "-0.249977111117893" },
         :high     => { :_theme => "4", :_tint => "-0.499984740745262" },
         :low      => { :_theme => "4", :_tint => "-0.499984740745262" }
       },
       {   # 20
         :series   => { :_theme => "5", :_tint => "0.39997558519241921" },
         :negative => { :_theme => "0", :_tint => "-0.499984740745262" },
         :markers  => { :_theme => "5", :_tint => "0.79998168889431442" },
         :first    => { :_theme => "5", :_tint => "-0.249977111117893" },
         :last     => { :_theme => "5", :_tint => "-0.249977111117893" },
         :high     => { :_theme => "5", :_tint => "-0.499984740745262" },
         :low      => { :_theme => "5", :_tint => "-0.499984740745262" }
       },
       {   # 21
         :series   => { :_theme => "6", :_tint => "0.39997558519241921" },
         :negative => { :_theme => "0", :_tint => "-0.499984740745262" },
         :markers  => { :_theme => "6", :_tint => "0.79998168889431442" },
         :first    => { :_theme => "6", :_tint => "-0.249977111117893" },
         :last     => { :_theme => "6", :_tint => "-0.249977111117893" },
         :high     => { :_theme => "6", :_tint => "-0.499984740745262" },
         :low      => { :_theme => "6", :_tint => "-0.499984740745262" }
       },
       {   # 22
         :series   => { :_theme => "7", :_tint => "0.39997558519241921" },
         :negative => { :_theme => "0", :_tint => "-0.499984740745262" },
         :markers  => { :_theme => "7", :_tint => "0.79998168889431442" },
         :first    => { :_theme => "7", :_tint => "-0.249977111117893" },
         :last     => { :_theme => "7", :_tint => "-0.249977111117893" },
         :high     => { :_theme => "7", :_tint => "-0.499984740745262" },
         :low      => { :_theme => "7", :_tint => "-0.499984740745262" }
       },
       {   # 23
         :series   => { :_theme => "8", :_tint => "0.39997558519241921" },
         :negative => { :_theme => "0", :_tint => "-0.499984740745262" },
         :markers  => { :_theme => "8", :_tint => "0.79998168889431442" },
         :first    => { :_theme => "8", :_tint => "-0.249977111117893" },
         :last     => { :_theme => "8", :_tint => "-0.249977111117893" },
         :high     => { :_theme => "8", :_tint => "-0.499984740745262" },
         :low      => { :_theme => "8", :_tint => "-0.499984740745262" }
       },
       {   # 24
         :series   => { :_theme => "9", :_tint => "0.39997558519241921" },
         :negative => { :_theme => "0", :_tint => "-0.499984740745262" },
         :markers  => { :_theme => "9", :_tint => "0.79998168889431442" },
         :first    => { :_theme => "9", :_tint => "-0.249977111117893" },
         :last     => { :_theme => "9", :_tint => "-0.249977111117893" },
         :high     => { :_theme => "9", :_tint => "-0.499984740745262" },
         :low      => { :_theme => "9", :_tint => "-0.499984740745262" }
       },
       {   # 25
         :series   => { :_theme => "1", :_tint => "0.499984740745262" },
         :negative => { :_theme => "1", :_tint => "0.249977111117893" },
         :markers  => { :_theme => "1", :_tint => "0.249977111117893" },
         :first    => { :_theme => "1", :_tint => "0.249977111117893" },
         :last     => { :_theme => "1", :_tint => "0.249977111117893" },
         :high     => { :_theme => "1", :_tint => "0.249977111117893" },
         :low      => { :_theme => "1", :_tint => "0.249977111117893" }
       },
       {   # 26
         :series   => { :_theme => "1", :_tint => "0.34998626667073579" },
         :negative => { :_theme => "0", :_tint => "-0.249977111117893" },
         :markers  => { :_theme => "0", :_tint => "-0.249977111117893" },
         :first    => { :_theme => "0", :_tint => "-0.249977111117893" },
         :last     => { :_theme => "0", :_tint => "-0.249977111117893" },
         :high     => { :_theme => "0", :_tint => "-0.249977111117893" },
         :low      => { :_theme => "0", :_tint => "-0.249977111117893" }
       },
       {   # 27
         :series   => { :_rgb => "FF323232" },
         :negative => { :_rgb => "FFD00000" },
         :markers  => { :_rgb => "FFD00000" },
         :first    => { :_rgb => "FFD00000" },
         :last     => { :_rgb => "FFD00000" },
         :high     => { :_rgb => "FFD00000" },
         :low      => { :_rgb => "FFD00000" }
       },
       {   # 28
         :series   => { :_rgb => "FF000000" },
         :negative => { :_rgb => "FF0070C0" },
         :markers  => { :_rgb => "FF0070C0" },
         :first    => { :_rgb => "FF0070C0" },
         :last     => { :_rgb => "FF0070C0" },
         :high     => { :_rgb => "FF0070C0" },
         :low      => { :_rgb => "FF0070C0" }
       },
       {   # 29
         :series   => { :_rgb => "FF376092" },
         :negative => { :_rgb => "FFD00000" },
         :markers  => { :_rgb => "FFD00000" },
         :first    => { :_rgb => "FFD00000" },
         :last     => { :_rgb => "FFD00000" },
         :high     => { :_rgb => "FFD00000" },
         :low      => { :_rgb => "FFD00000" }
       },
       {   # 30
         :series   => { :_rgb => "FF0070C0" },
         :negative => { :_rgb => "FF000000" },
         :markers  => { :_rgb => "FF000000" },
         :first    => { :_rgb => "FF000000" },
         :last     => { :_rgb => "FF000000" },
         :high     => { :_rgb => "FF000000" },
         :low      => { :_rgb => "FF000000" }
       },
       {   # 31
         :series   => { :_rgb => "FF5F5F5F" },
         :negative => { :_rgb => "FFFFB620" },
         :markers  => { :_rgb => "FFD70077" },
         :first    => { :_rgb => "FF5687C2" },
         :last     => { :_rgb => "FF359CEB" },
         :high     => { :_rgb => "FF56BE79" },
         :low      => { :_rgb => "FFFF5055" }
       },
       {   # 32
         :series   => { :_rgb => "FF5687C2" },
         :negative => { :_rgb => "FFFFB620" },
         :markers  => { :_rgb => "FFD70077" },
         :first    => { :_rgb => "FF777777" },
         :last     => { :_rgb => "FF359CEB" },
         :high     => { :_rgb => "FF56BE79" },
         :low      => { :_rgb => "FFFF5055" }
       },
       {   # 33
         :series   => { :_rgb => "FFC6EFCE" },
         :negative => { :_rgb => "FFFFC7CE" },
         :markers  => { :_rgb => "FF8CADD6" },
         :first    => { :_rgb => "FFFFDC47" },
         :last     => { :_rgb => "FFFFEB9C" },
         :high     => { :_rgb => "FF60D276" },
         :low      => { :_rgb => "FFFF5367" }
       },
       {   # 34
         :series   => { :_rgb => "FF00B050" },
         :negative => { :_rgb => "FFFF0000" },
         :markers  => { :_rgb => "FF0070C0" },
         :first    => { :_rgb => "FFFFC000" },
         :last     => { :_rgb => "FFFFC000" },
         :high     => { :_rgb => "FF00B050" },
         :low      => { :_rgb => "FFFF0000" }
       },
       {   # 35
         :series   => { :_theme => "3" },
         :negative => { :_theme => "9" },
         :markers  => { :_theme => "8" },
         :first    => { :_theme => "4" },
         :last     => { :_theme => "5" },
         :high     => { :_theme => "6" },
         :low      => { :_theme => "7" }
       },
       {   # 36
         :series   => { :_theme => "1" },
         :negative => { :_theme => "9" },
         :markers  => { :_theme => "8" },
         :first    => { :_theme => "4" },
         :last     => { :_theme => "5" },
         :high     => { :_theme => "6" },
         :low      => { :_theme => "7" }
       }
      ]
    end
  end
end
