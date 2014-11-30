# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/chart'

class TestWriteDLbls < Test::Unit::TestCase
  def setup
    @chart = Writexlsx::Chart.new('Line')
    @chart.instance_variable_set(
                                 :@label_positions,
                                 {
                                   'center'      => 'ctr',
                                   'right'       => 'r',
                                   'left'        => 'l',
                                   'top'         => 't',
                                   'above'       => 't',
                                   'bottom'      => 'b',
                                   'below'       => 'b',
                                   'inside_base' => 'inBase',
                                   'inside_end'  => 'inEnd',
                                   'outside_end' => 'outEnd',
                                   'best_fit'    => 'bestFit'
                                 }
                                 )
    @chart.instance_variable_set(:@label_position_default, '')
    @series = Writexlsx::Chart::Series.new(@chart)
  end

  def test_write_d_lbls_value_only
    expected = '<c:dLbls><c:showVal val="1"/></c:dLbls>'

    labels = @series.__send__(:labels_properties, :value => 1)
    @chart.__send__(:write_d_lbls, labels)

    result = chart_writer_string
    assert_equal(expected, result)
  end

  def test_write_d_lbls_series_name_only
    expected = '<c:dLbls><c:showSerName val="1"/></c:dLbls>'

    labels = @series.__send__(:labels_properties, :series_name => 1)
    @chart.__send__(:write_d_lbls, labels)

    result = chart_writer_string
    assert_equal(expected, result)
  end

  def test_write_d_lbls_category_only
    expected = '<c:dLbls><c:showCatName val="1"/></c:dLbls>'

    labels = @series.__send__(:labels_properties, :category => 1)
    @chart.__send__(:write_d_lbls, labels)

    result = chart_writer_string
    assert_equal(expected, result)
  end

  def test_write_d_lbls_value_category_and_series
    expected = '<c:dLbls><c:showVal val="1"/><c:showCatName val="1"/><c:showSerName val="1"/></c:dLbls>'

    labels = @series.__send__(:labels_properties,
                             {
                               :value       => 1,
                               :category    => 1,
                               :series_name => 1
                             }
                             )
    @chart.__send__(:write_d_lbls, labels)

    result = chart_writer_string
    assert_equal(expected, result)
  end

  def test_write_d_lbls_position_center
    expected = '<c:dLbls><c:dLblPos val="ctr"/><c:showVal val="1"/></c:dLbls>'

    labels = @series.__send__(:labels_properties,
                             {
                               :value    => 1,
                               :position => 'center'
                             }
                             )
    @chart.__send__(:write_d_lbls, labels)

    result = chart_writer_string
    assert_equal(expected, result)
  end

  def test_write_d_lbls_position_left
    expected = '<c:dLbls><c:dLblPos val="l"/><c:showVal val="1"/></c:dLbls>'

    labels = @series.__send__(:labels_properties,
                             {
                               :value    => 1,
                               :position => 'left'
                             }
                             )
    @chart.__send__(:write_d_lbls, labels)

    result = chart_writer_string
    assert_equal(expected, result)
  end

  def test_write_d_lbls_position_right
    expected = '<c:dLbls><c:dLblPos val="r"/><c:showVal val="1"/></c:dLbls>'

    labels = @series.__send__(:labels_properties,
                             {
                               :value    => 1,
                               :position => 'right'
                             }
                             )
    @chart.__send__(:write_d_lbls, labels)

    result = chart_writer_string
    assert_equal(expected, result)
  end

  def test_write_d_lbls_position_top
    expected = '<c:dLbls><c:dLblPos val="t"/><c:showVal val="1"/></c:dLbls>'

    labels = @series.__send__(:labels_properties,
                             {
                               :value    => 1,
                               :position => 'top'
                             }
                             )
    @chart.__send__(:write_d_lbls, labels)

    result = chart_writer_string
    assert_equal(expected, result)
  end

  def test_write_d_lbls_position_above
    expected = '<c:dLbls><c:dLblPos val="t"/><c:showVal val="1"/></c:dLbls>'

    labels = @series.__send__(:labels_properties,
                             {
                               :value    => 1,
                               :position => 'above'
                             }
                             )
    @chart.__send__(:write_d_lbls, labels)

    result = chart_writer_string
    assert_equal(expected, result)
  end

  def test_write_d_lbls_position_bottom
    expected = '<c:dLbls><c:dLblPos val="b"/><c:showVal val="1"/></c:dLbls>'

    labels = @series.__send__(:labels_properties,
                             {
                               :value    => 1,
                               :position => 'bottom'
                             }
                             )
    @chart.__send__(:write_d_lbls, labels)

    result = chart_writer_string
    assert_equal(expected, result)
  end

  def test_write_d_lbls_position_below
    expected = '<c:dLbls><c:dLblPos val="b"/><c:showVal val="1"/></c:dLbls>'

    labels = @series.__send__(:labels_properties,
                             {
                               :value    => 1,
                               :position => 'below'
                             }
                             )
    @chart.__send__(:write_d_lbls, labels)

    result = chart_writer_string
    assert_equal(expected, result)
  end

  def test_write_d_lbls_pie
    expected = '<c:dLbls><c:showVal val="1"/><c:showLeaderLines val="1"/></c:dLbls>'

    labels = @series.__send__(:labels_properties,
                             {
                               :value        => 1,
                               :leader_lines => 1
                             }
                             )
    @chart.__send__(:write_d_lbls, labels)

    result = chart_writer_string
    assert_equal(expected, result)
  end

  def test_write_d_lbls_pie_position_empty
    expected = '<c:dLbls><c:showVal val="1"/><c:showLeaderLines val="1"/></c:dLbls>'

    labels = @series.__send__(:labels_properties,
                             {
                               :value        => 1,
                               :leader_lines => 1,
                               :position     => ''
                             }
                             )
    @chart.__send__(:write_d_lbls, labels)

    result = chart_writer_string
    assert_equal(expected, result)
  end

  def test_write_d_lbls_pie_position_center
    expected = '<c:dLbls><c:dLblPos val="ctr"/><c:showVal val="1"/><c:showLeaderLines val="1"/></c:dLbls>'

    labels = @series.__send__(:labels_properties,
                             {
                               :value        => 1,
                               :leader_lines => 1,
                               :position     => 'center'
                             }
                             )
    @chart.__send__(:write_d_lbls, labels)

    result = chart_writer_string
    assert_equal(expected, result)
  end

  def test_write_d_lbls_pie_position_inside_end
    expected = '<c:dLbls><c:dLblPos val="inEnd"/><c:showVal val="1"/><c:showLeaderLines val="1"/></c:dLbls>'

    labels = @series.__send__(:labels_properties,
                             {
                               :value        => 1,
                               :leader_lines => 1,
                               :position     => 'inside_end'
                             }
                             )
    @chart.__send__(:write_d_lbls, labels)

    result = chart_writer_string
    assert_equal(expected, result)
  end

  def test_write_d_lbls_pie_position_outside_end
    expected = '<c:dLbls><c:dLblPos val="outEnd"/><c:showVal val="1"/><c:showLeaderLines val="1"/></c:dLbls>'

    labels = @series.__send__(:labels_properties,
                             {
                               :value        => 1,
                               :leader_lines => 1,
                               :position     => 'outside_end'
                             }
                             )
    @chart.__send__(:write_d_lbls, labels)

    result = chart_writer_string
    assert_equal(expected, result)
  end

  def test_write_d_lbls_pie_position_best_fit
    expected = '<c:dLbls><c:dLblPos val="bestFit"/><c:showVal val="1"/><c:showLeaderLines val="1"/></c:dLbls>'

    labels = @series.__send__(:labels_properties,
                             {
                               :value        => 1,
                               :leader_lines => 1,
                               :position     => 'best_fit'
                             }
                             )
    @chart.__send__(:write_d_lbls, labels)

    result = chart_writer_string
    assert_equal(expected, result)
  end

  def test_write_d_lbls_pie_percentage
    expected = '<c:dLbls><c:showPercent val="1"/><c:showLeaderLines val="1"/></c:dLbls>'

    labels = @series.__send__(:labels_properties,
                             {
                               :leader_lines => 1,
                               :percentage   => 1
                             }
                             )
    @chart.__send__(:write_d_lbls, labels)

    result = chart_writer_string
    assert_equal(expected, result)
  end

  def chart_writer_string
    @chart.instance_variable_get(:@writer).string
  end
end
