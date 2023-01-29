# -*- coding: utf-8 -*-
# frozen_string_literal: true

require 'singleton'

class ColName
  include Singleton

  def initialize
    @col_str_table = {}
  end

  def col_str(col)
    @col_str_table[col] = col_str_build(col) unless @col_str_table[col]
    @col_str_table[col]
  end

  private

  def col_str_build(col)
    # Change from 0-indexed to 1 indexed.
    col = col.to_i + 1
    col_str = ''

    while col > 0
      # Set remainder from 1 .. 26
      remainder = col % 26
      remainder = 26 if remainder == 0

      # Convert the remainder to a character. C-ishly.
      col_letter = ("A".ord + remainder - 1).chr

      # Accumulate the column letters, right to left.
      col_str = col_letter + col_str

      # Get the next order of magnitude.
      col = (col - 1) / 26
    end

    col_str
  end
end
