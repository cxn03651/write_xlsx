# -*- coding: utf-8 -*-

class Series
  attr_accessor :categories, :values, :name, :name_formula, :name_id
  attr_accessor :cat_data_id, :val_data_id, :line, :fill, :marker
  attr_accessor :trendline, :smooth, :labels, :invert_if_neg
  attr_accessor :x2_axis, :y2_axis, :error_bars, :points

  def ==(other)
    methods = %w[categories values name name_formula name_id
                 cat_data_id val_data_id
                 line fill marker trendline
                 smooth labels invert_if_neg x2_axis y2_axis error_bars points ]
    methods.each do |method|
      return false unless self.instance_variable_get("@#{method}") == other.instance_variable_get("@#{method}")
    end
    true
  end
end
