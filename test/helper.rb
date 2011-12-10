# -*- coding: utf-8 -*-
require 'rubygems'
require 'bundler'
begin
  Bundler.setup(:default, :development)
rescue Bundler::BundlerError => e
  $stderr.puts e.message
  $stderr.puts "Run `bundle install` to install missing gems"
  exit e.status_code
end
require 'test/unit'

$LOAD_PATH.unshift(File.dirname(__FILE__))
$LOAD_PATH.unshift(File.join(File.dirname(__FILE__), '..', 'lib'))
require 'write_xlsx'

class Test::Unit::TestCase
  def setup_dir_var
    @test_dir = File.dirname(__FILE__)
    @expected_dir = File.join(@test_dir, 'expected_dir')
    @result_dir   = File.join(@test_dir, 'result_dir')
    @perl_output  = File.join(@test_dir, 'perl_output')
  end

  def expected_to_array(lines)
    array = []
    lines.each_line do |line|
      str = line.chomp.sub(%r!/>$!, ' />').sub(/^\s+/, '')
      array << str unless str == ''
    end
    array
  end

  def got_to_array(xml_str)
    str = xml_str.gsub(/[\r\n]/, '')
    str.gsub(/>[ \t\r\n]*</, ">\t<").split(/\t/)
  end
end
