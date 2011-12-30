# -*- coding: utf-8 -*-
###############################################################################
#
# Chartsheet - A class for writing the Excel XLSX Chartsheet files.
#
# Used in conjunction with Excel::Writer::XLSX
#
# Copyright 2000-2011, John McNamara, jmcnamara@cpan.org
# Convert to ruby by Hideo NAKAMURA, cxn03651@msj.biglobe.ne.jp
#

require 'write_xlsx/worksheet'

module Writexlsx
  class Chartsheet < Worksheet
    include Writexlsx::Utility

    def initialize(workbook, index, name)
      super
      @drawing = 1
      @is_chartsheet     = true
      @chart             = nil
      @charts            = [1]
      @zoom_scale_normal = 0
      @oriantation       = 0
    end

    #
    # Assemble and write the XML file.
    #
    def assemble_xml_file
      return unless @writer
      write_xml_declaration

      # Write the root chartsheet element.
      write_chartsheet

      # Write the worksheet properties.
      write_sheet_pr

      # Write the sheet view properties.
      write_sheet_views

      # Write the sheetProtection element.
      write_sheet_protection

      # Write the printOptions element.
      write_print_options

      # Write the worksheet page_margins.
      write_page_margins

      # Write the worksheet page setup.
      write_page_setup

      # Write the headerFooter element.
      write_header_footer

      # Write the drawing element.
      write_drawings

      # Close the worksheet tag.
      @writer.end_tag('chartsheet')

      # Close the XML writer object and filehandle.
      @writer.crlf
      @writer.getOutput->close
    end

    def protect(password = '', options = {})
      @chart.protection = 1

      options[:sheet]     = 0
      options[:content]   = 1
      options[:scenarios] = 1

      protect(password, options)
    end


    ###############################################################################
    #
    # Encapsulated Chart methods.
    #
    ###############################################################################

    def add_series(*args)
      @chart.add_series(*args)
    end

    def set_x_axis(*args)
      @chart.set_x_axis(*args)
    end

    def set_y_axis(*args)
      @chart.set_y_axis(*args)
    end

    def set_title(*args)
      @chart.set_title(*args)
    end

    def set_legend(*args)
      @chart.set_legend(*args)
    end

    def set_plotarea(*args)
      @chart.set_plotarea(*args)
    end

    def set_chartarea(*args)
      @chart.set_chartarea(*args)
    end

    def set_style(*args)
      @chart.set_style(*args)
    end

    private

    #
    # Set up chart/drawings.
    #
    def prepare_chart(index, chart_id, drawing_id)
      drawing = Drawing.new
      @drawing = $drawing
      @drawing.orientation = @orientation

      @external_drawing_links << [ '/drawing', '../drawings/drawing' << drawing_id << '.xml' ]

      @drawing_links << [ '/chart', '../charts/chart' << chart_id << '.xml' ]
    end

    #
    # Write the <chartsheet> element. This is the root element of Chartsheet.
    #
    def write_chartsheet
      schema                 = 'http://schemas.openxmlformats.org/'
      xmlns                  = schema + 'spreadsheetml/2006/main'
      xmlns_r                = schema + 'officeDocument/2006/relationships'
      xmlns_mc               = schema + 'markup-compatibility/2006'
      xmlns_mv               = 'urn:schemas-microsoft-com:mac:vml'
      mc_ignorable           = 'mv'
      mc_preserve_attributes = 'mv:*'

      attributes = [
                    'xmlns',   xmlns,
                    'xmlns:r', xmlns_r
                   ]

      @writer.start_tag('chartsheet', attributes)
    end

    #
    # Write the <sheetPr> element for Sheet level properties.
    #
    def _write_sheet_pr

      attributes = []

      attributes << {'filterMode' => 1} if @filter_on

      if @fit_page || @tab_color
        @writer.start_tag('sheetPr', attributes)
        write_tab_color
        write_page_set_up_pr
        @writer.end_tag('sheetPr')
      else
        @writer.empty_tag('sheetPr', attributes)
      end
    end
  end
end
