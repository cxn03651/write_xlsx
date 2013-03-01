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

    attr_accessor :_type
    attr_accessor :_locations, :_ranges, :_count, :_high, :_low
    attr_accessor :_negative, :_first, :_last, :_markers, :_min, :_max
    attr_accessor :_axis, :_reverse, :_hidden, :_weight, :_empty
    attr_accessor :_date_axis, :_series_color, :_negative_color, :_markers_color
    attr_accessor :_first_color, :_last_color, :_high_color, :_low_color
    attr_accessor :_series_color, :_negative_color, :_markers_color
    attr_accessor :_first_color, :_last_color, :_high_color, :_low_color
    attr_accessor :_max, :_min, :_cust_max, :_cust_min, :_reverse

    def initialize
      @color = {}
    end

    def set_spark_color(user_color, palette_color)
      return unless palette_color

      instance_variable_set("@_#{user_color}", { :_rgb => palette_color })
    end
  end
end
