# -*- coding: utf-8 -*-
# frozen_string_literal: true

require 'bundler'

begin
  Bundler.setup(:default, :development)
rescue Bundler::BundlerError => e
  warn e.message
  warn "Run `bundle install` to install missing gems"
  exit e.status_code
end
require 'minitest/autorun'

$LOAD_PATH.unshift(File.dirname(__FILE__))
$LOAD_PATH.unshift(File.join(File.dirname(__FILE__), '..', 'lib'))
require 'write_xlsx'
require 'stringio'
require 'tempfile'

class Writexlsx::Workbook
  #
  # Set the default index for each format. This is mainly used for testing.
  #
  def set_default_xf_indices # :nodoc:
    # Delete the default url format.
    @formats.formats.delete_at(1)

    @formats.formats.each do |format|
      format.get_xf_index
    end
  end
end

class Minitest::Test
  def setup_dir_var
    @test_dir = File.dirname(__FILE__)
    @perl_output = File.join(@test_dir, 'perl_output')
    @regression_output = File.join(@test_dir, 'regression', 'xlsx_files')
    @io = StringIO.new(''.dup)
  end

  def expected_xlsx
    File.join(@regression_output, @xlsx)
  end

  def expected_to_array(lines)
    array = []
    lines.each_line do |line|
      str = line.chomp.sub(/^\s+/, '')
      array << str unless str == ''
    end
    array
  end

  def got_to_array(xml_str)
    xml_str.gsub(/[\r\n]/, '')
      .gsub(%r{ +/>}, ' />')
      .gsub(/>[ \t]*</, ">\t<")
      .split("\t")
  end

  def vml_str_to_array(vml_str)
    ret = ''
    vml_str.split(/[\r\n]+/).each do |vml|
      str = vml.sub(/^\s+/, '')     # Remove leading whitespace.
              .sub(/\s+$/, '')             # Remove trailing whitespace.
              .gsub("'", '"')             # Convert VMLs attribute quotes.
              .gsub(%r{([^ ])/>$}, '\1 />') # Add space before element end like XML::Writer.
              .sub(/"$/, '" ')        # Add space between attributes.
              .sub(/>$/, ">\n")       # Add newline after element end.
              .gsub("><", ">\n<")      # Split multiple elements.
      str.chomp! if str == "<x:Anchor>\n" # Put all of Anchor on one line.
      ret += str
    end
    ret.split("\n")
  end

  def entrys(xlsx)
    result = []
    Zip::File.foreach(xlsx) { |entry| result << entry }
    result
  end

  def compare_for_regression(ignore_members = nil, ignore_elements = nil)
    store_to_tempfile
    compare_xlsx(expected_xlsx, @tempfile.path, ignore_members, ignore_elements, true)
  end

  def store_to_tempfile
    @tempfile = Tempfile.open(@xlsx)
    @tempfile.binmode
    @tempfile.write(@io.string)
    @tempfile.close
  end

  def compare_xlsx_for_regression(exp_filename, got_filename, ignore_members = nil, ignore_elements = nil)
    compare_xlsx(exp_filename, got_filename, ignore_members, ignore_elements, true)
  end

  def compare_xlsx(exp_filename, got_filename, ignore_members = nil, ignore_elements = nil, regression = false)
    exp_members = filtered_members(exp_filename, ignore_members)
    got_members = filtered_members(got_filename, ignore_members)

    assert_equal(
      exp_members.map(&:name),
      got_members.map(&:name),
      "file members differs."
    )

    exp_members.zip(got_members).each do |exp_member, got_member|
      compare_member(
        exp_member,
        got_member,
        ignore_elements: ignore_elements,
        regression:      regression
      )
    end
  end

  def sort_rel_file_data(xml_array)
    header = xml_array.shift
    tail   = xml_array.pop
    xml_array.sort.unshift(header).push(tail)
  end

  #
  # Build worksheet XML from a block and return it as a string.
  #
  def worksheet_xml_string
    workbook  = WriteXLSX.new(StringIO.new(''.dup))
    worksheet = workbook.add_worksheet

    yield(workbook, worksheet)

    worksheet.assemble_xml_file
    worksheet.instance_variable_get(:@writer).string
  end

  #
  # Compare worksheet XML strings using the same normalization style as the
  # existing xlsx regression helpers.
  #
  def compare_worksheet_xml(expected_xml, actual_xml)
    exp_xml = got_to_array(expected_xml)
    got_xml = got_to_array(actual_xml)

    assert_equal(exp_xml, got_xml, 'worksheet xml differs.')
  end

  #
  # Assert that a worksheet XML string includes all lines in an expected
  # XML fragment after normalization.
  #
  def assert_worksheet_xml_includes(actual_xml, expected_fragment)
    got_xml = got_to_array(actual_xml)
    exp_xml = got_to_array(expected_fragment)

    exp_xml.each do |line|
      assert_includes(got_xml, line, "worksheet xml does not include: #{line}")
    end
  end

  #
  # Assert that a worksheet XML string does not include any lines in an
  # unexpected XML fragment after normalization.
  #
  def refute_worksheet_xml_includes(actual_xml, unexpected_fragment)
    got_xml = got_to_array(actual_xml)
    exp_xml = got_to_array(unexpected_fragment)

    exp_xml.each do |line|
      refute_includes(got_xml, line, "worksheet xml unexpectedly includes: #{line}")
    end
  end

  #
  # Extract the first conditionalFormatting element from worksheet XML.
  #
  def extract_conditional_formatting_xml(xml)
    xml[%r{<conditionalFormatting\b.*?</conditionalFormatting>}m]
  end

  private

  def filtered_members(filename, ignore_members)
    members = entrys(filename).sort_by(&:name)
    return members unless ignore_members

    members.reject { |member| ignore_members.include?(member.name) }
  end

  def compare_member(exp_member, got_member, ignore_elements:, regression:)
    exp_xml_str, got_xml_str = normalized_member_strings(
      exp_member,
      got_member,
      regression: regression
    )

    exp_xml = xml_array_for(exp_member.name, exp_xml_str)
    got_xml = xml_array_for(got_member.name, got_xml_str)

    if ignore_elements&.[](exp_member.name)
      exp_xml, got_xml = remove_ignored_elements(
        exp_xml,
        got_xml,
        ignore_elements[exp_member.name]
      )
    end

    if rel_file?(exp_member.name)
      exp_xml = sort_rel_file_data(exp_xml)
      got_xml = sort_rel_file_data(got_xml)
    end

    assert_equal(exp_xml, got_xml, "#{exp_member.name} differs.")
  end

  def normalized_member_strings(exp_member, got_member, regression:)
    got_str = got_member.get_input_stream.read
    exp_str = exp_member.get_input_stream.read

    ruby_19 do
      exp_str.force_encoding("ASCII-8BIT") if got_str.encoding == Encoding::ASCII_8BIT
    end

    got_xml_str = normalize_empty_tag_spacing(got_str)
    exp_xml_str = normalize_empty_tag_spacing(exp_str)

    [exp_xml_str, got_xml_str].tap do |exp_got|
      exp_got[0] = normalize_by_member_name(exp_member.name, exp_got[0], regression: regression, expected: true)
      exp_got[1] = normalize_by_member_name(got_member.name, exp_got[1], regression: regression, expected: false)
    end
  rescue StandardError
    p ruby_19 { got_str.encoding } || ruby_18 { got_str }
    p ruby_19 { exp_str.encoding } || ruby_18 { exp_str }
    raise
  end

  def normalize_empty_tag_spacing(str)
    str.gsub(%r{(\S)/>}, '\1 />')
  end

  def normalize_by_member_name(member_name, xml_str, regression:, expected:)
    case member_name
    when 'docProps/core.xml'
      normalize_core_xml(xml_str, regression: regression, expected: expected)
    when 'xl/workbook.xml'
      normalize_workbook_xml(xml_str)
    when %r{xl/worksheets/sheet\d.xml}
      normalize_worksheet_xml(xml_str, expected: expected)
    when %r{xl/charts/chart\d.xml}
      normalize_chart_xml(xml_str)
    else
      xml_str
    end
  end

  def normalize_core_xml(xml_str, regression:, expected:)
    xml_str = xml_str.gsub(/\d\d\d\d-\d\d-\d\dT\d\d:\d\d:\d\dZ/, '')

    if expected && regression
      xml_str = xml_str.gsub(/ ?John/, '')
    end

    xml_str
  end

  def normalize_workbook_xml(xml_str)
    xml_str
      .sub(/<workbookView[^>]*>/, '<workbookView/>')
      .sub(/<calcPr[^>]*>/, '<calcPr/>')
  end

  def normalize_worksheet_xml(xml_str, expected:)
    normalized = xml_str
                   .sub('horizontalDpi="200" ', '')
                   .sub('verticalDpi="200" ', '')

    if expected
      normalized = normalized
                     .sub(/(<pageSetup[^>]* )r:id="rId1"/, '\1')
                     .sub(%r{ +/>}, ' />')
    end

    normalized
  end

  def normalize_chart_xml(xml_str)
    xml_str.sub(/<c:pageMargins[^>]*>/, '<c:pageMargins/>')
  end

  def xml_array_for(member_name, xml_str)
    if member_name =~ /\.vml$/
      vml_str_to_array(xml_str)
    else
      got_to_array(xml_str)
    end
  end

  def remove_ignored_elements(exp_xml, got_xml, ignored_patterns)
    regex = Regexp.new(ignored_patterns.join('|'))

    [
      exp_xml.reject { |s| s =~ regex },
      got_xml.reject { |s| s =~ regex }
    ]
  end

  def rel_file?(member_name)
    member_name == '[Content_Types].xml' || member_name =~ /\.rels$/
  end
end
