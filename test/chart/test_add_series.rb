# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/chart'

class TestAddSeries < Test::Unit::TestCase
  def setup
    @chart = Writexlsx::Chart.new('Bar')
  end

  def test_add_series_only_values
    expected = {
      :_categories    => nil,
      :_values        => '=Sheet1!$A$1:$A$5',
      :_name          => nil,
      :_name_formula  => nil,
      :_name_id       => nil,
      :_cat_data_id   => nil,
      :_val_data_id   => 0,
      :_line          => { :_defined => 0 },
      :_fill          => { :_defined => 0 },
      :_marker        => nil,
      :_trendline     => nil,
      :_labels        => nil,
      :_invert_if_neg => nil,
      :_x2_axis       => nil,
      :_y2_axis       => nil,
      :_error_bars    => {
        :_x_error_bars => nil,
        :_y_error_bars => nil
      },
      :_points        => nil
    }

    @chart.add_series(:values => '=Sheet1!$A$1:$A$5')

    result = @chart.instance_variable_get(:@series).first
    assert_equal(expected, result)
  end

  def test_add_series_with_categories_and_values
    expected = [
                {
                  :_categories   => '=Sheet1!$A$1:$A$5',
                  :_values       => '=Sheet1!$B$1:$B$5',
                  :_name         => 'Text',
                  :_name_formula => nil,
                  :_name_id      => nil,
                  :_cat_data_id  => 0,
                  :_val_data_id  => 1,
                  :_line         => { :_defined => 0 },
                  :_fill         => { :_defined => 0 },
                  :_marker       => nil,
                  :_trendline    => nil,
                  :_labels        => nil,
                  :_invert_if_neg => nil,
                  :_x2_axis       => nil,
                  :_y2_axis       => nil,
                  :_error_bars    => {
                    :_x_error_bars => nil,
                    :_y_error_bars => nil
                  },
                  :_points        => nil
                }
               ]

    @chart.add_series(
                      :categories => '=Sheet1!$A$1:$A$5',
                      :values     => '=Sheet1!$B$1:$B$5',
                      :name       => 'Text'
                      )

    result = @chart.instance_variable_get(:@series)
    assert_equal(expected, result)
  end

  def test_add_series_only_values_checked_by_array
    expected = [
                {
                  :_categories   => nil,
                  :_values       => '=Sheet1!$A$1:$A$5',
                  :_name         => nil,
                  :_name_formula => nil,
                  :_name_id      => nil,
                  :_cat_data_id  => nil,
                  :_val_data_id  => 0,
                  :_line         => { :_defined => 0 },
                  :_fill         => { :_defined => 0 },
                  :_marker       => nil,
                  :_trendline    => nil,
                  :_labels       => nil,
                  :_invert_if_neg => nil,
                  :_x2_axis       => nil,
                  :_y2_axis       => nil,
                  :_error_bars    => {
                    :_x_error_bars => nil,
                    :_y_error_bars => nil
                  },
                  :_points        => nil
                }
               ]

    @chart.add_series(:values => ['Sheet1', 0, 4, 0, 0])

    result = @chart.instance_variable_get(:@series)
    assert_equal(expected, result)
  end

  def test_add_series_both_checked_by_array
    expected = {
      :_categories   => '=Sheet1!$A$1:$A$5',
      :_values       => '=Sheet1!$B$1:$B$5',
      :_name         => 'Text',
      :_name_formula => nil,
      :_name_id      => nil,
      :_cat_data_id  => 0,
      :_val_data_id  => 1,
      :_line         => { :_defined => 0 },
      :_fill         => { :_defined => 0 },
      :_marker       => nil,
      :_trendline    => nil,
      :_labels       => nil,
      :_invert_if_neg => nil,
      :_x2_axis       => nil,
      :_y2_axis       => nil,
      :_error_bars    => {
        :_x_error_bars => nil,
        :_y_error_bars => nil
      },
      :_points        => nil
   }

    @chart.add_series(
                      :categories => [ 'Sheet1', 0, 4, 0, 0 ],
                      :values     => [ 'Sheet1', 0, 4, 1, 1 ],
                      :name       => 'Text'
                      )

    result = @chart.instance_variable_get(:@series).first
    assert_equal(expected, result)
  end

  def test_add_series_secondary_axis
    expected = {
      :_categories    => '=Sheet1!$A$1:$A$5',
      :_values        => '=Sheet1!$B$1:$B$5',
      :_name          => 'Text',
      :_name_formula  => nil,
      :_name_id       => nil,
      :_cat_data_id   => 0,
      :_val_data_id   => 1,
      :_line          => { :_defined => 0 },
      :_fill          => { :_defined => 0 },
      :_marker        => nil,
      :_trendline     => nil,
      :_labels        => nil,
      :_invert_if_neg => nil,
      :_x2_axis       => 1,
      :_y2_axis       => 1,
      :_error_bars    => {
        :_x_error_bars => nil,
        :_y_error_bars => nil
      },
      :_points        => nil
    }

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
