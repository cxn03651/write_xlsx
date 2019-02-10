# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionProperties01 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_properties01
    @xlsx = 'properties01.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    workbook.set_properties(
                            :title    => 'This is an example spreadsheet',
                            :subject  => 'With document properties',
                            :author   => 'Someone',
                            :manager  => 'Dr. Heinz Doofenshmirtz',
                            :company  => 'of Wolves',
                            :category => 'Example spreadsheets',
                            :keywords => 'Sample, Example, Properties',
                            :comments => 'Created with Perl and Excel::Writer::XLSX',
                            :status   => 'Quo'
                            )

    worksheet.set_column('A:A', 70)
    worksheet.write('A1', "Select 'Office Button -> Prepare -> Properties' to see the file properties.")

    workbook.close
    compare_for_regression(
      nil,
      { 'xl/workbook.xml' => ['<workbookView'] }
    )
  end
end
