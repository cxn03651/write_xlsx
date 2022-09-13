# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionCondFormat19 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_cond_format19_with_double_quote
    create_xlsx('"X"')
    compare_for_regression(
      nil,
      { 'xl/workbook.xml' => ['<workbookView'] }
    )
  end

  def test_cond_format19_without_double_quote
    create_xlsx('X')
    compare_for_regression(
      nil,
      { 'xl/workbook.xml' => ['<workbookView'] }
    )
  end

  private

  def create_xlsx(value)
    @xlsx = 'cond_format19.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    format    = workbook.add_format(
      :color         => '#9C0006',
      :bg_color      => '#FFC7CE',
      :font_condense => 1,
      :font_extend   => 1
    )

    worksheet.write('A1', 10)
    worksheet.write('A2', 20)
    worksheet.write('A3', 30)
    worksheet.write('A4', 40)

    worksheet.conditional_formatting(
      'A1',
      {
        :type     => 'cell',
        :format   => format,
        :criteria => '==',
        :value    => value
      }
    )

    workbook.close
    compare_for_regression(
      nil,
      { 'xl/workbook.xml' => ['<workbookView'] }
    )
  end
end
