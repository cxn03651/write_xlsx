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

    attr_reader   :spark_color
    attr_writer   :_type
    attr_accessor :_locations, :_ranges, :_count, :_high, :_low
    attr_accessor :_negative, :_first, :_last, :_markers, :_min, :_max
    attr_accessor :_axis, :_reverse, :_hidden, :_weight, :_empty
    attr_accessor :_date_axis, :_series_color, :_negative_color, :_markers_color
    attr_accessor :_first_color, :_last_color, :_high_color, :_low_color

    def spark_color=(args)
      spark_color, color = args
      return unless color

      @spark_color ||= {}
      @spark_color[spark_color] = { :_rgb => color }
    end

    def [](attribute)
      instance_variable_get("@#{attribute}")
    end

    def []=(attribute, value)
      instance_variable_set("@#{attribute}", value)
    end
  end
end
