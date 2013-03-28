# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'
require 'stringio'

class TestTable13 < Test::Unit::TestCase
  def setup
    workbook = WriteXLSX.new(StringIO.new)
    @worksheet = workbook.add_worksheet
  end

  def test_table13_add_table_should_not_change_style_string
    style = 'Table Style Light 17'
    # Set the table properties.
    @worksheet.add_table('D4:I15', :style => style)

    # add_table should not change style string.
    assert_equal('Table Style Light 17', style)
  end

  def test_table13_add_table_should_not_change_formula_string
    formula = '=SUM(Table1[@[Column1]:[Column3]])'
    @worksheet.add_table(
                         'C2:F14',
                         {
                           :total_row => 1,
                           :columns   => [
                                          {:total_string => 'Total'},
                                          {},
                                          {},
                                          {
                                            :total_function => 'count',
                                            :format         => @format,
                                            :formula        => formula
                                          }
                                         ]
                         }
                         )

    # add_table should not change style string.
    assert_equal('=SUM(Table1[@[Column1]:[Column3]])', formula)
  end

  def test_table13_add_table_should_not_change_total_function_string
    total_function = 'std Dev'
    # Set the table properties.

    @worksheet.add_table(
                         'B2:K8',
                         {
                           :total_row => 1,
                           :columns => [
                                        {:total_string => 'Total'},
                                        {},
                                        {:total_function => 'Average'},
                                        {:total_function => 'COUNT'},
                                        {:total_function => 'count_nums'},
                                        {:total_function => 'max'},
                                        {:total_function => 'min'},
                                        {:total_function => 'sum'},
                                        {:total_function => total_function},
                                        {:total_function => 'var'}
                                       ]
                         }
                         )
    # add_table should not change total_function string.
    assert_equal('std Dev', total_function)
  end
end
