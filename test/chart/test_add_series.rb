# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/chart'

class TestAddSeries < Test::Unit::TestCase
  def setup
    @chart = Writexlsx::Chart.new('Bar')
  end

  def test_add_series_only_values
    series = Series.new
    series.categories      = nil
    series.values          = '=Sheet1!$A$1:$A$5'
    series.name            = nil
    series.name_formula   = nil
    series.name_id         = nil
    series.cat_data_id     = nil
    series.val_data_id    = 0
    series.line           = { :_defined => 0 }
    series.fill           = { :_defined => 0 }
    series.marker = nil
    series.trendline = nil
    series.smooth = nil
    series.labels = nil
    series.invert_if_neg = nil
    series.x2_axis = nil
    series.y2_axis = nil
    series.error_bars = {
        :_x_error_bars => nil,
        :_y_error_bars => nil
    }
    series.points = nil

    expected = series

    @chart.add_series(:values => '=Sheet1!$A$1:$A$5')

    result = @chart.instance_variable_get(:@series).first
    assert_equal(expected, result)
  end

  def test_add_series_with_categories_and_values
    series = Series.new
    series.categories     = '=Sheet1!$A$1:$A$5'
    series.values         = '=Sheet1!$B$1:$B$5'
    series.name           = 'Text'
    series.name_formula   = nil
    series.name_id        = nil
    series.cat_data_id    = 0
    series.val_data_id = 1
    series.line           = { :_defined => 0 }
    series.fill           = { :_defined => 0 }
    series.marker = nil
    series.trendline = nil
    series.smooth = nil
    series.labels = nil
    series.invert_if_neg = nil
    series.x2_axis = nil
    series.y2_axis = nil
    series.error_bars = {
        :_x_error_bars => nil,
        :_y_error_bars => nil
    }
    series.points = nil
    expected = [ series ]

    @chart.add_series(
                      :categories => '=Sheet1!$A$1:$A$5',
                      :values     => '=Sheet1!$B$1:$B$5',
                      :name       => 'Text'
                      )

    result = @chart.instance_variable_get(:@series)
    assert_equal(expected, result)
  end

  def test_add_series_only_values_checked_by_array
    series = Series.new
    series.categories     = nil
    series.values         = '=Sheet1!$A$1:$A$5'
    series.name           = nil
    series.name_formula   = nil
    series.name_id        = nil
    series.cat_data_id    = nil
    series.val_data_id = 0
    series.line           = { :_defined => 0 }
    series.fill           = { :_defined => 0 }
    series.marker = nil
    series.trendline = nil
    series.smooth = nil
    series.labels = nil
    series.invert_if_neg = nil
    series.x2_axis = nil
    series.y2_axis = nil
    series.error_bars = {
        :_x_error_bars => nil,
        :_y_error_bars => nil
    }
    series.points = nil
    expected = [ series ]

    @chart.add_series(:values => ['Sheet1', 0, 4, 0, 0])

    result = @chart.instance_variable_get(:@series)
    assert_equal(expected, result)
  end

  def test_add_series_both_checked_by_array
    series = Series.new
    series.categories     = '=Sheet1!$A$1:$A$5'
    series.values         = '=Sheet1!$B$1:$B$5'
    series.name           = 'Text'
    series.name_formula   = nil
    series.name_id        = nil
    series.cat_data_id   = 0
    series.val_data_id = 1
    series.line           = { :_defined => 0 }
    series.fill           = { :_defined => 0 }
    series.marker = nil
    series.trendline = nil
    series.smooth = nil
    series.labels = nil
    series.invert_if_neg = nil
    series.x2_axis = nil
    series.y2_axis = nil
    series.error_bars = {
        :_x_error_bars => nil,
        :_y_error_bars => nil
    }
    series.points = nil
    expected = series

    @chart.add_series(
                      :categories => [ 'Sheet1', 0, 4, 0, 0 ],
                      :values     => [ 'Sheet1', 0, 4, 1, 1 ],
                      :name       => 'Text'
                      )

    result = @chart.instance_variable_get(:@series).first
    assert_equal(expected, result)
  end

  def test_add_series_secondary_axis
    series = Series.new
    series.categories     = '=Sheet1!$A$1:$A$5'
    series.values         = '=Sheet1!$B$1:$B$5'
    series.name           = 'Text'
    series.name_formula   = nil
    series.name_id        = nil
    series.cat_data_id    = 0
    series.val_data_id = 1
    series.line           = { :_defined => 0 }
    series.fill           = { :_defined => 0 }
    series.marker = nil
    series.trendline = nil
    series.smooth = nil
    series.labels = nil
    series.invert_if_neg = nil
    series.x2_axis = 1
    series.y2_axis = 1
    series.error_bars = {
        :_x_error_bars => nil,
        :_y_error_bars => nil
    }
    series.points = nil
    expected = series

    @chart.add_series(
                      :categories => [ 'Sheet1', 0, 4, 0, 0 ],
                      :values     => [ 'Sheet1', 0, 4, 1, 1 ],
                      :name       => 'Text',
                      :x2_axis    => 1,
                      :y2_axis    => 1
                      )

    result = @chart.instance_variable_get(:@series).first
    assert_equal(expected, result)
  end
end
