# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'
require 'stringio'

class TestParseFilterExpression < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  def test_parse_filter_expression
    [
      [
          'x =  2000',
          [2, 2000],
      ],

      [
          'x == 2000',
          [2, 2000],
      ],

      [
          'x =~ 2000',
          [2, 2000],
      ],

      [
          'x eq 2000',
          [2, 2000],
      ],

      [
          'x <> 2000',
          [5, 2000],
      ],

      [
          'x != 2000',
          [5, 2000],
      ],

      [
          'x ne 2000',
          [5, 2000],
      ],

      [
          'x !~ 2000',
          [5, 2000],
      ],

      [
          'x >  2000',
          [4, 2000],
      ],

      [
          'x <  2000',
          [1, 2000],
      ],

      [
          'x >= 2000',
          [6, 2000],
      ],

      [
          'x <= 2000',
          [3, 2000],
      ],

      [
          'x >  2000 and x <  5000',
          [4,  2000, 0, 1, 5000],
      ],

      [
          'x >  2000 &&  x <  5000',
          [4,  2000, 0, 1, 5000],
      ],

      [
          'x >  2000 or  x <  5000',
          [4,  2000, 1, 1, 5000],
      ],

      [
          'x >  2000 ||  x <  5000',
          [4,  2000, 1, 1, 5000],
      ],

      [
          'x =  Blanks',
          [2, 'blanks'],
      ],

      [
          'x =  NonBlanks',
          [5, ' '],
      ],

      [
          'x <> Blanks',
          [5, ' '],
      ],

      [
          'x <> NonBlanks',
          [2, 'blanks'],
      ],

      [
          'Top 10 Items',
          [30, 10],
      ],

      [
          'Top 20 %',
          [31, 20],
      ],

      [
          'Bottom 5 Items',
          [32, 5],
      ],

      [
          'Bottom 101 %',
          [33, 101],
      ]
    ].each do |test|
      expected = test[1]
      tokens   = @worksheet.__send__('extract_filter_tokens', test[0])
      result   = @worksheet.__send__('parse_filter_expression', test[0], tokens)

      testname = test[0] || 'none'

      assert_equal(expected, result, testname)
    end
  end
end
