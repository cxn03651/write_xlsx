# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/chart'

class TestAddSeries < Test::Unit::TestCase
  def setup
    @chart = Writexlsx::Chart.new('Bar')
  end

  def test_add_series_only_values
    series = Writexlsx::Chart::Series.new(@chart)
    series.instance_variable_set(:@categories, nil)
    series.instance_variable_set(:@values, '=Sheet1!$A$1:$A$5')
    series.instance_variable_set(:@name, nil)
    series.instance_variable_set(:@name_formula, nil)
    series.instance_variable_set(:@name_id, nil)
    series.instance_variable_set(:@cat_data_id, nil)
    series.instance_variable_set(:@val_data_id, 0)
    series.instance_variable_set(:@line, { :_defined => 0 })
    series.instance_variable_set(:@fill, { :_defined => 0 })
    series.instance_variable_set(:@marker, nil)
    series.instance_variable_set(:@trendline, nil)
    series.instance_variable_set(:@smooth, nil)
    series.instance_variable_set(:@labels, nil)
    series.instance_variable_set(:@invert_if_neg, nil)
    series.instance_variable_set(:@x2_axis, nil)
    series.instance_variable_set(:@y2_axis, nil)
    series.instance_variable_set(:@error_bars, {
        :_x_error_bars => nil,
        :_y_error_bars => nil
    })
    series.instance_variable_set(:@points, nil)

    expected = series

    @chart.add_series(:values => '=Sheet1!$A$1:$A$5')

    result = @chart.instance_variable_get(:@series).first
    assert_equal(expected, result)
  end

  def test_add_series_with_categories_and_values
    series = Writexlsx::Chart::Series.new(@chart)
    series.instance_variable_set(:@categories, '=Sheet1!$A$1:$A$5')
    series.instance_variable_set(:@values, '=Sheet1!$B$1:$B$5')
    series.instance_variable_set(:@name, 'Text')
    series.instance_variable_set(:@name_formula, nil)
    series.instance_variable_set(:@name_id       , nil)
    series.instance_variable_set(:@cat_data_id   , 0)
    series.instance_variable_set(:@val_data_id, 1)
    series.instance_variable_set(:@line          , { :_defined => 0 })
    series.instance_variable_set(:@fill          , { :_defined => 0 })
    series.instance_variable_set(:@marker, nil)
    series.instance_variable_set(:@trendline, nil)
    series.instance_variable_set(:@smooth, nil)
    series.instance_variable_set(:@labels, nil)
    series.instance_variable_set(:@invert_if_neg, nil)
    series.instance_variable_set(:@x2_axis, nil)
    series.instance_variable_set(:@y2_axis, nil)
    series.instance_variable_set(:@error_bars, {
        :_x_error_bars => nil,
        :_y_error_bars => nil
    })
    series.instance_variable_set(:@points, nil)
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
    series = Writexlsx::Chart::Series.new(@chart)
    series.instance_variable_set(:@categories, nil)
    series.instance_variable_set(:@values, '=Sheet1!$A$1:$A$5')
    series.instance_variable_set(:@name, nil)
    series.instance_variable_set(:@name_formula, nil)
    series.instance_variable_set(:@name_id       , nil)
    series.instance_variable_set(:@cat_data_id   , nil)
    series.instance_variable_set(:@val_data_id, 0)
    series.instance_variable_set(:@line          , { :_defined => 0 })
    series.instance_variable_set(:@fill          , { :_defined => 0 })
    series.instance_variable_set(:@marker, nil)
    series.instance_variable_set(:@trendline, nil)
    series.instance_variable_set(:@smooth, nil)
    series.instance_variable_set(:@labels, nil)
    series.instance_variable_set(:@invert_if_neg, nil)
    series.instance_variable_set(:@x2_axis, nil)
    series.instance_variable_set(:@y2_axis, nil)
    series.instance_variable_set(:@error_bars, {
        :_x_error_bars => nil,
        :_y_error_bars => nil
    })
    series.instance_variable_set(:@points, nil)
    expected = [ series ]

    @chart.add_series(:values => ['Sheet1', 0, 4, 0, 0])

    result = @chart.instance_variable_get(:@series)
    assert_equal(expected, result)
  end

  def test_add_series_both_checked_by_array
    series = Writexlsx::Chart::Series.new(@chart)
    series.instance_variable_set(:@categories, '=Sheet1!$A$1:$A$5')
    series.instance_variable_set(:@values, '=Sheet1!$B$1:$B$5')
    series.instance_variable_set(:@name, 'Text')
    series.instance_variable_set(:@name_formula, nil)
    series.instance_variable_set(:@name_id       , nil)
    series.instance_variable_set(:@cat_data_id  , 0)
    series.instance_variable_set(:@val_data_id, 1)
    series.instance_variable_set(:@line          , { :_defined => 0 })
    series.instance_variable_set(:@fill          , { :_defined => 0 })
    series.instance_variable_set(:@marker, nil)
    series.instance_variable_set(:@trendline, nil)
    series.instance_variable_set(:@smooth, nil)
    series.instance_variable_set(:@labels, nil)
    series.instance_variable_set(:@invert_if_neg, nil)
    series.instance_variable_set(:@x2_axis, nil)
    series.instance_variable_set(:@y2_axis, nil)
    series.instance_variable_set(:@error_bars, {
        :_x_error_bars => nil,
        :_y_error_bars => nil
    })
    series.instance_variable_set(:@points, nil)
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
    series = Writexlsx::Chart::Series.new(@chart)
    series.instance_variable_set(:@categories,  '=Sheet1!$A$1:$A$5')
    series.instance_variable_set(:@values, '=Sheet1!$B$1:$B$5')
    series.instance_variable_set(:@name, 'Text')
    series.instance_variable_set(:@name_formula, nil)
    series.instance_variable_set(:@name_id, nil)
    series.instance_variable_set(:@cat_data_id, 0)
    series.instance_variable_set(:@val_data_id, 1)
    series.instance_variable_set(:@line          , { :_defined => 0 })
    series.instance_variable_set(:@fill          , { :_defined => 0 })
    series.instance_variable_set(:@marker, nil)
    series.instance_variable_set(:@trendline, nil)
    series.instance_variable_set(:@smooth, nil)
    series.instance_variable_set(:@labels, nil)
    series.instance_variable_set(:@invert_if_neg, nil)
    series.instance_variable_set(:@x2_axis, 1)
    series.instance_variable_set(:@y2_axis, 1)
    series.instance_variable_set(:@error_bars,
                                 {
                                   :_x_error_bars => nil,
                                   :_y_error_bars => nil
                                 })
    series.instance_variable_set(:@points, nil)
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
