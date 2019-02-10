# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionChartClustered01 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_chart_clustered01
    @xlsx = 'chart_clustered01.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    chart     = workbook.add_chart(:type => 'column', :embedded => 1)

    # For testing, copy the randomly generated axis ids in the target xlsx file.
    chart.instance_variable_set(:@axis_ids, [45886080, 45928832])

    data = [
      ['Types',  'Sub Type',  'Value 1', 'Value 2', 'Value 3'],
      ['Type 1', 'Sub Type A', 5000,      8000,      6000],
      ['',       'Sub Type B', 2000,      3000,      4000],
      ['',       'Sub Type C', 250,       1000,      2000],
      ['Type 2', 'Sub Type D', 6000,      6000,      6500],
      ['',       'Sub Type E', 500,       300,       200]
    ]

    cat_data = [
      ['Type 1',     nil,          nil,          'Type 2',     nil],
      ['Sub Type A', 'Sub Type B', 'Sub Type C', 'Sub Type D', 'Sub Type E']
    ]

    worksheet.write_col('A1', data)

    chart.add_series(
      :name            => '=Sheet1!$C$1',
      :categories      => '=Sheet1!$A$2:$B$6',
      :values          => '=Sheet1!$C$2:$C$6',
      :categories_data => cat_data
    );

    chart.add_series(
      :name       => '=Sheet1!$D$1',
      :categories => '=Sheet1!$A$2:$B$6',
      :values     => '=Sheet1!$D$2:$D$6'
    )

    chart.add_series(
      :name       => '=Sheet1!$E$1',
      :categories => '=Sheet1!$A$2:$B$6',
      :values     => '=Sheet1!$E$2:$E$6'
    )

    worksheet.insert_chart('E9', chart)

    workbook.close
    compare_for_regression(
      [],
      {'xl/charts/chart1.xml' => [ '<c:pageMargins' ]}
    )
  end
end
