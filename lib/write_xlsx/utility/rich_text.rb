# -*- coding: utf-8 -*-
# frozen_string_literal: true

module Writexlsx
  module Utility
    module RichText
      def underline_attributes(underline)
        case underline
        when 2
          [%w[val double]]
        when 33
          [%w[val singleAccounting]]
        when 34
          [%w[val doubleAccounting]]
        else
          []    # Default to single underline.
        end
      end

      #
      # Convert user defined font values into private hash values.
      #
      def convert_font_args(params)
        return unless params

        font = params_to_font(params)

        # Convert font size units.
        font[:_size] *= 100 if font[:_size] && font[:_size] != 0

        # Convert rotation into 60,000ths of a degree.
        font[:_rotation] = 60_000 * font[:_rotation].to_i if ptrue?(font[:_rotation])

        font
      end

      def params_to_font(params)
        {
          _name:         params[:name],
          _color:        params[:color],
          _size:         params[:size],
          _bold:         params[:bold],
          _italic:       params[:italic],
          _underline:    params[:underline],
          _pitch_family: params[:pitch_family],
          _charset:      params[:charset],
          _baseline:     params[:baseline] || 0,
          _rotation:     params[:rotation]
        }
      end

      #
      # Get the font style attributes from a font hash.
      #
      def get_font_style_attributes(font)
        return [] unless font
        return [] unless font.respond_to?(:[])

        attributes = []
        attributes << ['sz', font[:_size]]      if ptrue?(font[:_size])
        attributes << ['b',  font[:_bold]]      if font[:_bold]
        attributes << ['i',  font[:_italic]]    if font[:_italic]
        attributes << %w[u sng]                 if font[:_underline]

        # Turn off baseline when testing fonts that don't have it.
        attributes << ['baseline', font[:_baseline]] if font[:_baseline] != -1
        attributes
      end

      #
      # Get the font latin attributes from a font hash.
      #
      def get_font_latin_attributes(font)
        return [] unless font
        return [] unless font.respond_to?(:[])

        attributes = []
        attributes << ['typeface', font[:_name]]            if ptrue?(font[:_name])
        attributes << ['pitchFamily', font[:_pitch_family]] if font[:_pitch_family]
        attributes << ['charset', font[:_charset]]          if font[:_charset]

        attributes
      end

      #
      # Write the <c:txPr> element.
      #
      def write_tx_pr(font, is_y_axis = nil) # :nodoc:
        rotation = nil
        rotation = font[:_rotation] if font && font.respond_to?(:[]) && font[:_rotation]
        @writer.tag_elements('c:txPr') do
          # Write the a:bodyPr element.
          write_a_body_pr(rotation, is_y_axis)
          # Write the a:lstStyle element.
          write_a_lst_style
          # Write the a:p element.
          write_a_p_formula(font)
        end
      end

      #
      # Write the <a:bodyPr> element.
      #
      def write_a_body_pr(rot, is_y_axis = nil) # :nodoc:
        rot = -5400000 if !rot && ptrue?(is_y_axis)
        attributes = []
        if rot
          if rot == 16_200_000
            # 270 deg/stacked angle.
            attributes << ['rot',  0]
            attributes << %w[vert wordArtVert]
          elsif rot == 16_260_000
            # 271 deg/stacked angle.
            attributes << ['rot',  0]
            attributes << %w[vert eaVert]
          else
            attributes << ['rot',  rot]
            attributes << %w[vert horz]
          end
        end

        @writer.empty_tag('a:bodyPr', attributes)
      end

      #
      # Write the <a:lstStyle> element.
      #
      def write_a_lst_style # :nodoc:
        @writer.empty_tag('a:lstStyle')
      end

      #
      # Write the <a:p> element for formula titles.
      #
      def write_a_p_formula(font = nil) # :nodoc:
        @writer.tag_elements('a:p') do
          # Write the a:pPr element.
          write_a_p_pr_formula(font)
          # Write the a:endParaRPr element.
          write_a_end_para_rpr
        end
      end

      #
      # Write the <a:pPr> element for formula titles.
      #
      def write_a_p_pr_formula(font) # :nodoc:
        @writer.tag_elements('a:pPr') { write_a_def_rpr(font) }
      end

      #
      # Write the <a:defRPr> element.
      #
      def write_a_def_rpr(font = nil) # :nodoc:
        write_def_rpr_r_pr_common(
          font,
          get_font_style_attributes(font),
          'a:defRPr'
        )
      end

      def write_def_rpr_r_pr_common(font, style_attributes, tag)  # :nodoc:
        latin_attributes = get_font_latin_attributes(font)
        has_color = ptrue?(font) && ptrue?(font[:_color])

        if !latin_attributes.empty? || has_color
          @writer.tag_elements(tag, style_attributes) do
            write_a_solid_fill(color: font[:_color]) if has_color
            write_a_latin(latin_attributes) unless latin_attributes.empty?
          end
        else
          @writer.empty_tag(tag, style_attributes)
        end
      end

      #
      # Write the <a:endParaRPr> element.
      #
      def write_a_end_para_rpr # :nodoc:
        @writer.empty_tag('a:endParaRPr', [%w[lang en-US]])
      end
    end
  end
end
