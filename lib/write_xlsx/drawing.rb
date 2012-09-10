# -*- coding: utf-8 -*-
require 'write_xlsx/package/xml_writer_simple.rb'

module Writexlsx
  class Drawing
    attr_writer :embedded, :orientation

    def initialize
      @writer = Package::XMLWriterSimple.new
      @drawings    = []
      @embedded    = false
      @orientation = false
    end

    def xml_str
      @writer.string
    end

    def set_xml_writer(filename)
      @writer.set_xml_writer(filename)
    end

    #
    # Assemble and write the XML file.
    #
    def assemble_xml_file
      @writer.xml_decl

      # Write the xdr:wsDr element.
      write_drawing_workspace

      if @embedded
          index = 0
          @drawings.each do |dimensions|
            # Write the xdr:twoCellAnchor element.
            index += 1

            write_two_cell_anchor(index, *(dimensions.flatten))
          end
      else
        index = 0

        # Write the xdr:absoluteAnchor element.
        index += 1
        write_absolute_anchor(index)
      end

      @writer.end_tag('xdr:wsDr')
      @writer.crlf
      @writer.close
    end

    #
    # Add a chart or image sub object to the drawing.
    #
    def add_drawing_object(*args)
      @drawings << args
    end

    private

    #
    # Write the <xdr:wsDr> element.
    #
    def write_drawing_workspace
      schema    = 'http://schemas.openxmlformats.org/drawingml/'
      xmlns_xdr = "#{schema}2006/spreadsheetDrawing"
      xmlns_a   = "#{schema}2006/main"

      attributes = [
          'xmlns:xdr', xmlns_xdr,
          'xmlns:a',   xmlns_a
      ]

      @writer.start_tag('xdr:wsDr', attributes)
    end

    #
    # Write the <xdr:twoCellAnchor> element.
    #
    def write_two_cell_anchor(*args)
      index, type, col_from, row_from, col_from_offset, row_from_offset,
      col_to, row_to, col_to_offset, row_to_offset, col_absolute, row_absolute,
      width, height, description = args

      attributes      = []

      # Add attribute for images.
      attributes << :editAs << 'oneCell' if type == 2

      @writer.tag_elements('xdr:twoCellAnchor', attributes) do
        # Write the xdr:from element.
        write_from(col_from, row_from, col_from_offset, row_from_offset)
        # Write the xdr:from element.
        write_to(col_to, row_to, col_to_offset, row_to_offset)

        if type == 1
          # Write the xdr:graphicFrame element for charts.
          write_graphic_frame(index)
        else
          # Write the xdr:pic element.
          write_pic(index, col_absolute, row_absolute, width, height, description)
        end

        # Write the xdr:clientData element.
        write_client_data
      end
    end

    #
    # Write the <xdr:absoluteAnchor> element.
    #
    def write_absolute_anchor(index)
      @writer.start_tag('xdr:absoluteAnchor')

      # Different co-ordinates for horizonatal (= 0) and vertical (= 1).
      if !@orientation || @orientation == 0

        # Write the xdr:pos element.
        write_pos(0, 0)

        # Write the xdr:ext element.
        write_ext(9308969, 6078325)
      else
        # Write the xdr:pos element.
        write_pos(0, -47625)

        # Write the xdr:ext element.
        write_ext(6162675, 6124575)
      end

      # Write the xdr:graphicFrame element.
      write_graphic_frame(index)

      # Write the xdr:clientData element.
      write_client_data

      @writer.end_tag('xdr:absoluteAnchor')
    end

      #
    # Write the <xdr:from> element.
    #
    def write_from(col, row, col_offset, row_offset)
      @writer.tag_elements('xdr:from') do
        # Write the xdr:col element.
        write_col(col)
        # Write the xdr:colOff element.
        write_col_off(col_offset)
        # Write the xdr:row element.
        write_row(row)
        # Write the xdr:rowOff element.
        write_row_off(row_offset)
      end
    end

    #
    # Write the <xdr:to> element.
    #
    def write_to(col, row, col_offset, row_offset)
      @writer.tag_elements('xdr:to') do
        # Write the xdr:col element.
        write_col(col)
        # Write the xdr:colOff element.
        write_col_off(col_offset)
        # Write the xdr:row element.
        write_row(row)
        # Write the xdr:rowOff element.
        write_row_off(row_offset)
      end
    end

    #
    # Write the <xdr:col> element.
    #
    def write_col(data)
      @writer.data_element('xdr:col', data)
    end

    #
    # Write the <xdr:colOff> element.
    #
    def write_col_off(data)
      @writer.data_element('xdr:colOff', data)
    end


    #
    # Write the <xdr:row> element.
    #
    def write_row(data)
      @writer.data_element('xdr:row', data)
    end


    #
    # Write the <xdr:rowOff> element.
    #
    def write_row_off(data)
      @writer.data_element('xdr:rowOff', data)
    end

    #
    # Write the <xdr:pos> element.
    #
    def write_pos(x, y)
      attributes = [
                    'x', x,
                    'y', y
                   ]

      @writer.empty_tag('xdr:pos', attributes)
    end

    #
    # Write the <xdr:ext> element.
    #
    def write_ext(cx, cy)
      attributes = [
                    'cx', cx,
                    'cy', cy
                   ]

      @writer.empty_tag('xdr:ext', attributes)
    end

    #
    # Write the <xdr:graphicFrame> element.
    #
    def write_graphic_frame(index)
      macro  = ''

      attributes = ['macro', macro]

      @writer.tag_elements('xdr:graphicFrame', attributes) do
        # Write the xdr:nvGraphicFramePr element.
        write_nv_graphic_frame_pr(index)
        # Write the xdr:xfrm element.
        write_xfrm
        # Write the a:graphic element.
        write_atag_graphic(index)
      end
    end

    #
    # Write the <xdr:nvGraphicFramePr> element.
    #
    def write_nv_graphic_frame_pr(index)
      @writer.tag_elements('xdr:nvGraphicFramePr') do
        # Write the xdr:cNvPr element.
        write_c_nv_pr( index + 1, "Chart #{index}" )
        # Write the xdr:cNvGraphicFramePr element.
        write_c_nv_graphic_frame_pr
      end
    end

    #
    # Write the <xdr:cNvPr> element.
    #
    def write_c_nv_pr(id, name, descr = nil)
      attributes = [
          'id',   id,
          'name', name
      ]

      # Add description attribute for images.
      attributes << 'descr' << descr if descr

      @writer.empty_tag('xdr:cNvPr', attributes)
    end


    #
    # Write the <xdr:cNvGraphicFramePr> element.
    #
    def write_c_nv_graphic_frame_pr
      if @embedded
        @writer.empty_tag('xdr:cNvGraphicFramePr')
      else
        @writer.tag_elements('xdr:cNvGraphicFramePr') do
          # Write the a:graphicFrameLocks element.
          write_a_graphic_frame_locks
        end
      end
    end

    #
    # Write the <a:graphicFrameLocks> element.
    #
    def write_a_graphic_frame_locks
      no_grp = 1

      attributes = ['noGrp', no_grp ]

      @writer.empty_tag('a:graphicFrameLocks', attributes)
    end

    #
    # Write the <xdr:xfrm> element.
    #
    def write_xfrm
      @writer.tag_elements('xdr:xfrm') do
        # Write the xfrmOffset element.
        write_xfrm_offset
        # Write the xfrmOffset element.
        write_xfrm_extension
      end
    end

    #
    # Write the <a:off> xfrm sub-element.
    #
    def write_xfrm_offset
      x    = 0
      y    = 0

      attributes = [
        'x', x,
        'y', y
      ]

      @writer.empty_tag('a:off', attributes)
    end

    #
    # Write the <a:ext> xfrm sub-element.
    #
    def write_xfrm_extension
      x    = 0
      y    = 0

      attributes = [
        'cx', x,
        'cy', y
      ]

      @writer.empty_tag('a:ext', attributes)
    end

    #
    # Write the <a:graphic> element.
    #
    def write_atag_graphic(index)
      @writer.tag_elements('a:graphic') do
        # Write the a:graphicData element.
        write_atag_graphic_data(index)
      end
    end

    #
    # Write the <a:graphicData> element.
    #
    def write_atag_graphic_data(index)
      uri   = 'http://schemas.openxmlformats.org/drawingml/2006/chart'

      attributes = ['uri', uri]

      @writer.tag_elements('a:graphicData', attributes) do
        # Write the c:chart element.
        write_c_chart("rId#{index}")
      end
    end

    #
    # Write the <c:chart> element.
    #
    def write_c_chart(r_id)
      schema  = 'http://schemas.openxmlformats.org/'
      xmlns_c = "#{schema}drawingml/2006/chart"
      xmlns_r = "#{schema}officeDocument/2006/relationships"


      attributes = [
        'xmlns:c', xmlns_c,
        'xmlns:r', xmlns_r,
        'r:id',    r_id
      ]

      @writer.empty_tag('c:chart', attributes)
    end

    #
    # Write the <xdr:clientData> element.
    #
    def write_client_data
      @writer.empty_tag('xdr:clientData')
    end

    #
    # Write the <xdr:pic> element.
    #
    def write_pic(index, col_absolute, row_absolute, width, height, description)
      @writer.tag_elements('xdr:pic') do
        # Write the xdr:nvPicPr element.
        write_nv_pic_pr(index, description)
        # Write the xdr:blipFill element.
        write_blip_fill(index)
        # Write the xdr:spPr element.
        write_sp_pr(col_absolute, row_absolute, width, height)
      end
    end

    #
    # Write the <xdr:nvPicPr> element.
    #
    def write_nv_pic_pr(index, description)
      @writer.tag_elements('xdr:nvPicPr') do
        # Write the xdr:cNvPr element.
        write_c_nv_pr( index + 1, "Picture #{index}", description )
        # Write the xdr:cNvPicPr element.
        write_c_nv_pic_pr
      end
    end

    #
    # Write the <xdr:cNvPicPr> element.
    #
    def write_c_nv_pic_pr
      @writer.tag_elements('xdr:cNvPicPr') do
        # Write the a:picLocks element.
        write_a_pic_locks
      end
    end

    #
    # Write the <a:picLocks> element.
    #
    def write_a_pic_locks
      no_change_aspect = 1

      attributes = ['noChangeAspect', no_change_aspect]

      @writer.empty_tag('a:picLocks', attributes)
    end

    #
    # Write the <xdr:blipFill> element.
    #
    def write_blip_fill(index)
      @writer.tag_elements('xdr:blipFill') do
        # Write the a:blip element.
        write_a_blip(index)
        # Write the a:stretch element.
        write_a_stretch
      end
    end

    #
    # Write the <a:blip> element.
    #
    def write_a_blip(index)
      schema  = 'http://schemas.openxmlformats.org/officeDocument/'
      xmlns_r = "#{schema}2006/relationships"
      r_embed = "rId#{index}"

      attributes = [
        'xmlns:r', xmlns_r,
        'r:embed', r_embed
      ]

      @writer.empty_tag('a:blip', attributes)
    end

    #
    # Write the <a:stretch> element.
    #
    def write_a_stretch
      @writer.tag_elements('a:stretch') do
        # Write the a:fillRect element.
        write_a_fill_rect
      end
    end

    #
    # Write the <a:fillRect> element.
    #
    def write_a_fill_rect
      @writer.empty_tag('a:fillRect')
    end

    #
    # Write the <xdr:spPr> element.
    #
    def write_sp_pr(col_absolute, row_absolute, width, height)
      @writer.tag_elements('xdr:spPr') do
        # Write the a:xfrm element.
        write_a_xfrm(col_absolute, row_absolute, width, height)
        # Write the a:prstGeom element.
        write_a_prst_geom
      end
    end

    #
    # Write the <a:xfrm> element.
    #
    def write_a_xfrm(col_absolute, row_absolute, width, height)
      @writer.tag_elements('a:xfrm') do
        # Write the a:off element.
        write_a_off( col_absolute, row_absolute )
        # Write the a:ext element.
        write_a_ext( width, height )
      end
    end

    #
    # Write the <a:off> element.
    #
    def write_a_off(x, y)
      attributes = [
        'x', x,
        'y', y
      ]

      @writer.empty_tag('a:off', attributes)
    end


    #
    # Write the <a:ext> element.
    #
    def write_a_ext(cx, cy)
      attributes = [
        'cx', cx,
        'cy', cy
      ]

      @writer.empty_tag('a:ext', attributes)
    end

    #
    # Write the <a:prstGeom> element.
    #
    def write_a_prst_geom
      prst = 'rect'

      attributes = ['prst', prst]

      @writer.tag_elements('a:prstGeom', attributes) do
        # Write the a:avLst element.
        write_a_av_lst
      end
    end

    #
    # Write the <a:avLst> element.
    #
    def write_a_av_lst
      @writer.empty_tag('a:avLst')
    end
  end
end
