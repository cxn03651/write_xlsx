# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'
require 'stringio'

class TestWriteDataValidation01 < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  def test_write_data_validations_gt_zero
    @worksheet.data_validation('A1',
      :validate => 'integer',
      :criteria => '>',
      :value    => 0
    )
    @worksheet.__send__('write_data_validations')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<dataValidations count="1"><dataValidation type="whole" operator="greaterThan" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A1"><formula1>0</formula1></dataValidation></dataValidations>'
    assert_equal(expected, result)
  end

  def test_write_data_validations_gt_zero_with_options_turned_off
    @worksheet.data_validation('A1',
      :validate     => 'integer',
      :criteria     => '>',
      :value        => 0,
      :ignore_blank => 0,
      :show_input   => 0,
      :show_error   => 0
    )
    @worksheet.__send__('write_data_validations')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<dataValidations count="1"><dataValidation type="whole" operator="greaterThan" sqref="A1"><formula1>0</formula1></dataValidation></dataValidations>'
    assert_equal(expected, result)
  end

  def test_write_data_validations_gt_E3
    @worksheet.data_validation('A2',
      :validate => 'integer',
      :criteria => '>',
      :value    => 'E3'
    )
    @worksheet.__send__('write_data_validations')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<dataValidations count="1"><dataValidation type="whole" operator="greaterThan" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A2"><formula1>E3</formula1></dataValidation></dataValidations>'
    assert_equal(expected, result)
  end

  def test_write_data_validations_decimal_between_01_05
    @worksheet.data_validation('A3',
      :validate => 'decimal',
      :criteria => 'between',
      :minimum  => 0.1,
      :maximum  => 0.5
    )
    @worksheet.__send__('write_data_validations')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<dataValidations count="1"><dataValidation type="decimal" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A3"><formula1>0.1</formula1><formula2>0.5</formula2></dataValidation></dataValidations>'
    assert_equal(expected, result)
  end

  def test_write_data_validations_list_array
    @worksheet.data_validation('A4',
      :validate => 'list',
      :source   => [ 'open', 'high', 'close' ]
    )
    @worksheet.__send__('write_data_validations')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<dataValidations count="1"><dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A4"><formula1>"open,high,close"</formula1></dataValidation></dataValidations>'
    assert_equal(expected, result)
  end

  def test_write_data_validations_list_reference
    @worksheet.data_validation('A5',
      :validate => 'list',
      :source   => '=$E$4:$G$4'
    )
    @worksheet.__send__('write_data_validations')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<dataValidations count="1"><dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A5"><formula1>$E$4:$G$4</formula1></dataValidation></dataValidations>'
    assert_equal(expected, result)
  end

  def test_write_data_validations_list_date_between
    @worksheet.data_validation('A6',
      :validate => 'date',
      :criteria => 'between',
      :minimum  => '2008-01-01T',
      :maximum  => '2008-12-12T'
    )
    @worksheet.__send__('write_data_validations')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<dataValidations count="1"><dataValidation type="date" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A6"><formula1>39448</formula1><formula2>39794</formula2></dataValidation></dataValidations>'
    assert_equal(expected, result)
  end

  def test_write_data_validations_between_with_title_and_message
    @worksheet.data_validation('A7',
      :validate      => 'integer',
      :criteria      => 'between',
      :minimum       => 1,
      :maximum       => 100,
      :input_title   => 'Enter an integer:',
      :input_message => 'between 1 and 100'
    )
    @worksheet.__send__('write_data_validations')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<dataValidations count="1"><dataValidation type="whole" allowBlank="1" showInputMessage="1" showErrorMessage="1" promptTitle="Enter an integer:" prompt="between 1 and 100" sqref="A7"><formula1>1</formula1><formula2>100</formula2></dataValidation></dataValidations>'
    assert_equal(expected, result)
  end
end
