# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionDefaultRow05 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_default_row05
    @xlsx = 'default_row05.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.set_default_row(24, 1)

    worksheet.write('A1',  'Foo')
    worksheet.write('A10', 'Bar')
    worksheet.write('A20', 'Baz')

    # for my $row (1 .. 8, 10 .. 19) {
    (1..19).to_a.reject { |x| x == 9 }.each do |row|
      worksheet.set_row(row, 24)
    end
    
    workbook.close
    compare_for_regression
  end
end
