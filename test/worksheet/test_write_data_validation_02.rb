# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'
require 'stringio'

class TestWriteDataValidation02 < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  def test_write_data_validations_between_1_and_10
    @worksheet.data_validation('B5',
      :validate      => 'integer',
      :criteria      => 'between',
      :minimum       => 1,
      :maximum       => 10
    )
    @worksheet.__send__('write_data_validations')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<dataValidations count="1"><dataValidation type="whole" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>'
    assert_equal(expected, result)
  end

  def test_write_data_validations_not_between_1_and_10
    @worksheet.data_validation('B5',
      :validate      => 'integer',
      :criteria      => 'not between',
      :minimum       => 1,
      :maximum       => 10
    )
    @worksheet.__send__('write_data_validations')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<dataValidations count="1"><dataValidation type="whole" operator="notBetween" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>'
    assert_equal(expected, result)
  end

  def test_write_data_validations_integer_equal_1
    ['equal to', '=', '=='].each do |operator|
      workbook = WriteXLSX.new(StringIO.new)
      worksheet = workbook.add_worksheet('')
      worksheet.data_validation('B5',
        :validate      => 'integer',
        :criteria      => operator,
        :value         => 1
      )
      worksheet.__send__('write_data_validations')
      result = worksheet.instance_variable_get(:@writer).string
      expected = '<dataValidations count="1"><dataValidation type="whole" operator="equal" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>1</formula1></dataValidation></dataValidations>'
      assert_equal(expected, result)
    end
  end

  def test_write_data_validations_integer_not_equal_1
    ['not equal to', '<>', '!='].each do |operator|
      workbook = WriteXLSX.new(StringIO.new)
      worksheet = workbook.add_worksheet('')
      worksheet.data_validation('B5',
        :validate      => 'integer',
        :criteria      => operator,
        :value         => 1
      )
      worksheet.__send__('write_data_validations')
      result = worksheet.instance_variable_get(:@writer).string
      expected = '<dataValidations count="1"><dataValidation type="whole" operator="notEqual" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>1</formula1></dataValidation></dataValidations>'
      assert_equal(expected, result)
    end
  end

  def test_write_data_validations_integer_greater_than_1
    ['greater than', '>'].each do |operator|
      workbook = WriteXLSX.new(StringIO.new)
      worksheet = workbook.add_worksheet('')
      worksheet.data_validation('B5',
        :validate      => 'integer',
        :criteria      => operator,
        :value         => 1
      )
      worksheet.__send__('write_data_validations')
      result = worksheet.instance_variable_get(:@writer).string
      expected = '<dataValidations count="1"><dataValidation type="whole" operator="greaterThan" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>1</formula1></dataValidation></dataValidations>'
      assert_equal(expected, result)
    end
  end

  def test_write_data_validations_integer_less_than_1
    ['less than', '<'].each do |operator|
      workbook = WriteXLSX.new(StringIO.new)
      worksheet = workbook.add_worksheet('')
      worksheet.data_validation('B5',
        :validate      => 'integer',
        :criteria      => operator,
        :value         => 1
      )
      worksheet.__send__('write_data_validations')
      result = worksheet.instance_variable_get(:@writer).string
      expected = '<dataValidations count="1"><dataValidation type="whole" operator="lessThan" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>1</formula1></dataValidation></dataValidations>'
      assert_equal(expected, result)
    end
  end

  def test_write_data_validations_integer_greater_than_or_equal_to_1
    ['greater than or equal to', '>='].each do |operator|
      workbook = WriteXLSX.new(StringIO.new)
      worksheet = workbook.add_worksheet('')
      worksheet.data_validation('B5',
        :validate      => 'integer',
        :criteria      => operator,
        :value         => 1
      )
      worksheet.__send__('write_data_validations')
      result = worksheet.instance_variable_get(:@writer).string
      expected = '<dataValidations count="1"><dataValidation type="whole" operator="greaterThanOrEqual" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>1</formula1></dataValidation></dataValidations>'
      assert_equal(expected, result)
    end
  end

  def test_write_data_validations_integer_less_than_or_equal_to_1
    ['less than or equal to', '<='].each do |operator|
      workbook = WriteXLSX.new(StringIO.new)
      worksheet = workbook.add_worksheet('')
      worksheet.data_validation('B5',
        :validate      => 'integer',
        :criteria      => operator,
        :value         => 1
      )
      worksheet.__send__('write_data_validations')
      result = worksheet.instance_variable_get(:@writer).string
      expected = '<dataValidations count="1"><dataValidation type="whole" operator="lessThanOrEqual" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>1</formula1></dataValidation></dataValidations>'
      assert_equal(expected, result)
    end
  end

  def test_write_data_validations_integer_between_1_and_10_not_ignore_blank
    @worksheet.data_validation('B5',
      :validate      => 'integer',
      :criteria      => 'between',
      :minimum       => 1,
      :maximum       => 10,
      :ignore_blank  => 0
    )
    @worksheet.__send__('write_data_validations')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<dataValidations count="1"><dataValidation type="whole" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>'
    assert_equal(expected, result)
  end

  def test_write_data_validations_integer_between_1_and_10_error_type_warning
    @worksheet.data_validation('B5',
      :validate      => 'integer',
      :criteria      => 'between',
      :minimum       => 1,
      :maximum       => 10,
      :error_type    => 'warning'
    )
    @worksheet.__send__('write_data_validations')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<dataValidations count="1"><dataValidation type="whole" errorStyle="warning" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>'
    assert_equal(expected, result)
  end

  def test_write_data_validations_integer_between_1_and_10_error_type_information
    @worksheet.data_validation('B5',
      :validate      => 'integer',
      :criteria      => 'between',
      :minimum       => 1,
      :maximum       => 10,
      :error_type    => 'information'
    )
    @worksheet.__send__('write_data_validations')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<dataValidations count="1"><dataValidation type="whole" errorStyle="information" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>'
    assert_equal(expected, result)
  end

  def test_write_data_validations_integer_between_1_and_10_with_input_title
    @worksheet.data_validation('B5',
      :validate      => 'integer',
      :criteria      => 'between',
      :minimum       => 1,
      :maximum       => 10,
      :input_title   => 'Input title January'
    )
    @worksheet.__send__('write_data_validations')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<dataValidations count="1"><dataValidation type="whole" allowBlank="1" showInputMessage="1" showErrorMessage="1" promptTitle="Input title January" sqref="B5"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>'
    assert_equal(expected, result)
  end

  def test_write_data_validations_integer_between_1_and_10_with_input_title_and_input_message
    @worksheet.data_validation('B5',
      :validate      => 'integer',
      :criteria      => 'between',
      :minimum       => 1,
      :maximum       => 10,
      :input_title   => 'Input title January',
      :input_message => 'Input message February'
    )
    @worksheet.__send__('write_data_validations')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<dataValidations count="1"><dataValidation type="whole" allowBlank="1" showInputMessage="1" showErrorMessage="1" promptTitle="Input title January" prompt="Input message February" sqref="B5"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>'
    assert_equal(expected, result)
  end

  def test_write_data_validations_integer_between_1_and_10_with_input_title_and_input_message_and_error_title
    @worksheet.data_validation('B5',
      :validate      => 'integer',
      :criteria      => 'between',
      :minimum       => 1,
      :maximum       => 10,
      :input_title   => 'Input title January',
      :input_message => 'Input message February',
      :error_title   => 'Error title March'
    )
    @worksheet.__send__('write_data_validations')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<dataValidations count="1"><dataValidation type="whole" allowBlank="1" showInputMessage="1" showErrorMessage="1" errorTitle="Error title March" promptTitle="Input title January" prompt="Input message February" sqref="B5"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>'
    assert_equal(expected, result)
  end

  def test_write_data_validations_integer_between_1_and_10_with_input_title_and_input_message_and_error_title_and_error_message
    @worksheet.data_validation('B5',
      :validate      => 'integer',
      :criteria      => 'between',
      :minimum       => 1,
      :maximum       => 10,
      :input_title   => 'Input title January',
      :input_message => 'Input message February',
      :error_title   => 'Error title March',
      :error_message => 'Error message April'
    )
    @worksheet.__send__('write_data_validations')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<dataValidations count="1"><dataValidation type="whole" allowBlank="1" showInputMessage="1" showErrorMessage="1" errorTitle="Error title March" error="Error message April" promptTitle="Input title January" prompt="Input message February" sqref="B5"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>'
    assert_equal(expected, result)
  end

  def test_write_data_validations_integer_between_1_and_10_with_input_title_and_input_message_and_error_title_and_error_message_and_show_input
    @worksheet.data_validation('B5',
      :validate      => 'integer',
      :criteria      => 'between',
      :minimum       => 1,
      :maximum       => 10,
      :input_title   => 'Input title January',
      :input_message => 'Input message February',
      :error_title   => 'Error title March',
      :error_message => 'Error message April',
      :show_input    => 0
    )
    @worksheet.__send__('write_data_validations')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<dataValidations count="1"><dataValidation type="whole" allowBlank="1" showErrorMessage="1" errorTitle="Error title March" error="Error message April" promptTitle="Input title January" prompt="Input message February" sqref="B5"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>'
    assert_equal(expected, result)
  end

  def test_write_data_validations_integer_between_1_and_10_with_input_title_and_input_message_and_error_title_and_error_message_and_show_input_and_show_error
    @worksheet.data_validation('B5',
      :validate      => 'integer',
      :criteria      => 'between',
      :minimum       => 1,
      :maximum       => 10,
      :input_title   => 'Input title January',
      :input_message => 'Input message February',
      :error_title   => 'Error title March',
      :error_message => 'Error message April',
      :show_input    => 0,
      :show_error    => 0
    )
    @worksheet.__send__('write_data_validations')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<dataValidations count="1"><dataValidation type="whole" allowBlank="1" errorTitle="Error title March" error="Error message April" promptTitle="Input title January" prompt="Input message February" sqref="B5"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>'
    assert_equal(expected, result)
  end

  def test_write_data_validation_validate_any
    @worksheet.data_validation('B5', :validate => 'any')
    @worksheet.__send__('write_data_validations')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = ''
    assert_equal(expected, result)
  end

  def test_write_data_validation_decimal_equal_to_12345
    @worksheet.data_validation('B5',
      :validate => 'decimal',
      :criteria => '==',
      :value    => 1.2345
    )
    @worksheet.__send__('write_data_validations')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<dataValidations count="1"><dataValidation type="decimal" operator="equal" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>1.2345</formula1></dataValidation></dataValidations>'
    assert_equal(expected, result)
  end

  def test_write_data_validation_list_a_bb_ccc
    @worksheet.data_validation('B5',
      :validate => 'list',
      :source   => ['a', 'bb', 'ccc']
    )
    @worksheet.__send__('write_data_validations')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<dataValidations count="1"><dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>"a,bb,ccc"</formula1></dataValidation></dataValidations>'
    assert_equal(expected, result)
  end

  def test_write_data_validation_list_a_bb_ccc_without_dropdown
    @worksheet.data_validation('B5',
      :validate => 'list',
      :source   => ['a', 'bb', 'ccc'],
      :dropdown => 0
    )
    @worksheet.__send__('write_data_validations')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<dataValidations count="1"><dataValidation type="list" allowBlank="1" showDropDown="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>"a,bb,ccc"</formula1></dataValidation></dataValidations>'
    assert_equal(expected, result)
  end

  def test_write_data_validation_list_absolute_range
    @worksheet.data_validation(
      'A1:A1',
      :validate => 'list',
      :source   => '=$D$1:$D$5'
    )
    @worksheet.__send__('write_data_validations')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<dataValidations count="1"><dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A1"><formula1>$D$1:$D$5</formula1></dataValidation></dataValidations>'
    assert_equal(expected, result)
  end

  def test_write_data_validation_date_equal_to_39653
    @worksheet.data_validation(
      'B5',
      :validate => 'date',
      :criteria => '==',
      :value    => 39653
    )
    @worksheet.__send__('write_data_validations')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<dataValidations count="1"><dataValidation type="date" operator="equal" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>39653</formula1></dataValidation></dataValidations>'
    assert_equal(expected, result)
  end

  def test_write_data_validation_date_equal_to_2008_07_24T
    @worksheet.data_validation(
      'B5',
      :validate => 'date',
      :criteria => '==',
      :value    => '2008-07-24T'
    )
    @worksheet.__send__('write_data_validations')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<dataValidations count="1"><dataValidation type="date" operator="equal" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>39653</formula1></dataValidation></dataValidations>'
    assert_equal(expected, result)
  end

  def test_write_data_validation_date_between_range
    @worksheet.data_validation(
      'B5',
      :validate => 'date',
      :criteria => 'between',
      :minimum  => '2008-01-01T',
      :maximum  => '2008-12-12T'
    )
    @worksheet.__send__('write_data_validations')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<dataValidations count="1"><dataValidation type="date" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>39448</formula1><formula2>39794</formula2></dataValidation></dataValidations>'
    assert_equal(expected, result)
  end

  def test_write_data_validation_time_equal_to_05
    @worksheet.data_validation(
      'B5:B5',
      :validate => 'time',
      :criteria => '==',
      :value    => 0.5
    )
    @worksheet.__send__('write_data_validations')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<dataValidations count="1"><dataValidation type="time" operator="equal" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>0.5</formula1></dataValidation></dataValidations>'
    assert_equal(expected, result)
  end

  def test_write_data_validation_time_equal_to_1200
    @worksheet.data_validation(
      'B5',
      :validate => 'time',
      :criteria => '==',
      :value    => 'T12:00:00'
    )
    @worksheet.__send__('write_data_validations')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<dataValidations count="1"><dataValidation type="time" operator="equal" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>0.5</formula1></dataValidation></dataValidations>'
    assert_equal(expected, result)
  end

  def test_write_data_validation_custom_equal_to_10
    @worksheet.data_validation(
      'B5',
      :validate => 'custom',
      :criteria => '==',
      :value    => 10
    )
    @worksheet.__send__('write_data_validations')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<dataValidations count="1"><dataValidation type="custom" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>10</formula1></dataValidation></dataValidations>'
    assert_equal(expected, result)
  end

  def test_write_data_validations_between_1_and_10_A1_cell
    @worksheet.data_validation(
      'B5',
      :validate      => 'integer',
      :criteria      => 'between',
      :minimum       => 1,
      :maximum       => 10
    )
    @worksheet.__send__('write_data_validations')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<dataValidations count="1"><dataValidation type="whole" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>'
    assert_equal(expected, result)
  end

  def test_write_data_validations_between_1_and_10_A1_range
    @worksheet.data_validation(
      'B5:B10',
      :validate      => 'integer',
      :criteria      => 'between',
      :minimum       => 1,
      :maximum       => 10
    )
    @worksheet.__send__('write_data_validations')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<dataValidations count="1"><dataValidation type="whole" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5:B10"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>'
    assert_equal(expected, result)
  end

  def test_write_data_validations_between_1_and_10_row_col_cell
    @worksheet.data_validation(
      4, 1,
      :validate      => 'integer',
      :criteria      => 'between',
      :minimum       => 1,
      :maximum       => 10
    )
    @worksheet.__send__('write_data_validations')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<dataValidations count="1"><dataValidation type="whole" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>'
    assert_equal(expected, result)
  end

  def test_write_data_validations_between_1_and_10_row_col_range
    @worksheet.data_validation(
      4, 1, 9, 1,
      :validate      => 'integer',
      :criteria      => 'between',
      :minimum       => 1,
      :maximum       => 10
    )
    @worksheet.__send__('write_data_validations')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<dataValidations count="1"><dataValidation type="whole" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5:B10"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>'
    assert_equal(expected, result)
  end

  def test_write_data_validations_multiple_style_cells
    @worksheet.data_validation(
      4, 1,
      :validate      => 'integer',
      :criteria      => 'between',
      :minimum       => 1,
      :maximum       => 10,
      :other_cells => [ [ 4, 3, 4, 3 ] ]
    )
    @worksheet.__send__('write_data_validations')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<dataValidations count="1"><dataValidation type="whole" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5 D5"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>'
    assert_equal(expected, result)
  end

  def test_write_data_validations_multiple_style_cells_2
    @worksheet.data_validation(
      4, 1,
      :validate      => 'integer',
      :criteria      => 'between',
      :minimum       => 1,
      :maximum       => 10,
      :other_cells => [ [ 6, 1, 6, 1 ], [ 8, 1, 8, 1 ] ]
    )
    @worksheet.__send__('write_data_validations')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<dataValidations count="1"><dataValidation type="whole" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5 B7 B9"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>'
    assert_equal(expected, result)
  end

  def test_write_data_validations_multiple_style_cells_3
    @worksheet.data_validation(
      4, 1, 8, 1,
      :validate      => 'integer',
      :criteria      => 'between',
      :minimum       => 1,
      :maximum       => 10,
      :other_cells => [ [ 3, 3, 3, 3 ] ]
    )
    @worksheet.__send__('write_data_validations')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<dataValidations count="1"><dataValidation type="whole" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5:B9 D4"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>'
    assert_equal(expected, result)
  end

  def test_write_data_validations_multiple_validation
    @worksheet.data_validation(
      'B5',
      :validate      => 'integer',
      :criteria      => '>',
      :value         => 10
    )
    @worksheet.data_validation(
      'C10',
      :validate      => 'integer',
      :criteria      => '<',
      :value         => 10
    )
    @worksheet.__send__('write_data_validations')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<dataValidations count="2"><dataValidation type="whole" operator="greaterThan" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>10</formula1></dataValidation><dataValidation type="whole" operator="lessThan" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="C10"><formula1>10</formula1></dataValidation></dataValidations>'
    assert_equal(expected, result)
  end
end
