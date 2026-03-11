# -*- coding: utf-8 -*-
# frozen_string_literal: true

module Writexlsx
  class Worksheet
    module DrawingXmlWriter
      #
      # Write the <drawing> elements.
      #
      def write_drawings # :nodoc:
        increment_rel_id_and_write_r_id('drawing') if drawings?
      end

      #
      # Write the <legacyDrawing> element.
      #
      def write_legacy_drawing # :nodoc:
        increment_rel_id_and_write_r_id('legacyDrawing') if has_vml?
      end

      #
      # Write the <legacyDrawingHF> element.
      #
      def write_legacy_drawing_hf # :nodoc:
        return unless has_header_vml?

        # Increment the relationship id for any drawings or comments.
        @rel_count += 1

        attributes = [['r:id', "rId#{@rel_count}"]]
        @writer.empty_tag('legacyDrawingHF', attributes)
      end

      #
      # Write the <picture> element.
      #
      def write_picture
        return unless background_image

        # Increment the relationship id.
        @rel_count += 1
        id = @rel_count

        attributes = [['r:id', "rId#{id}"]]

        @writer.empty_tag('picture', attributes)
      end
    end
  end
end
