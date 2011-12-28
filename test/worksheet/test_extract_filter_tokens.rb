# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'
require 'stringio'

class TestExtractFilterTokens < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  def test_extract_filter_tokens
    [
      [
          nil,
          [],
      ],

      [
          '',
          [],
      ],

      [
          '0 <  2000',
          [0, '<', 2000],
      ],

      [
          'x <  2000',
          ['x', '<', 2000],
      ],

      [
          'x >  2000',
          ['x', '>', 2000],
      ],

      [
          'x == 2000',
          ['x', '==', 2000],
      ],

      [
          'x >  2000 and x <  5000',
          ['x', '>',  2000, 'and', 'x', '<', 5000],
      ],

      [
          'x = "foo"',
          ['x', '=', 'foo'],
      ],

      [
          'x = foo',
          ['x', '=', 'foo'],
      ],

      [
          'x = "foo bar"',
          ['x', '=', 'foo bar'],
      ],

      [
          'x = "foo "" bar"',
          ['x', '=', 'foo " bar'],
      ],

      [
          'x = "foo bar" or x = "bar foo"',
          ['x', '=', 'foo bar', 'or', 'x', '=', 'bar foo'],
      ],

      [
          'x = "foo "" bar" or x = "bar "" foo"',
          ['x', '=', 'foo " bar', 'or', 'x', '=', 'bar " foo'],
      ],

      [
          'x = """"""""',
          ['x', '=', '"""'],
      ],

      [
          'x = Blanks',
          ['x', '=', 'Blanks'],
      ],

      [
          'x = NonBlanks',
          ['x', '=', 'NonBlanks'],
      ],

      [
          'top 10 %',
          ['top', 10, '%'],
      ],

      [
          'top 10 items',
          ['top', 10, 'items'],
      ]
    ].each do |test|
      expected = test[1]
      result = @worksheet.__send__('extract_filter_tokens', test[0])
      assert_equal(expected, result)
    end
  end
end
