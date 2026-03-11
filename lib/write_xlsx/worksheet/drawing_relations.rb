# -*- coding: utf-8 -*-
# frozen_string_literal: true

module Writexlsx
  class Worksheet
    module DrawingRelations
      ###############################################################################
      #
      # DrawingRelations
      #
      # Manages drawing relationships and external linkage information.
      #
      # Responsibilities:
      # - Relationship ID allocation for drawings and VML objects
      # - Tracking drawing and VML relationship mappings
      # - Managing external links for drawings, comments, tables, and media
      # - Providing relationship data for rels file generation
      #
      # This module handles *how drawing resources are linked together*.
      # It does not prepare drawing objects or write XML output.
      #
      ###############################################################################
      #
      # Get the index used to address a drawing rel link.
      #
      def drawing_rel_index(target = nil)
        if !target
          # Undefined values for drawings like charts will always be unique.
          @drawing_rels_id += 1
        elsif ptrue?(@drawing_rels[target])
          @drawing_rels[target]
        else
          @drawing_rels_id += 1
          @drawing_rels[target] = @drawing_rels_id
        end
      end

      #
      # Get the index used to address a vml_drawing rel link.
      #
      def get_vml_drawing_rel_index(target)
        if @vml_drawing_rels[target]
          @vml_drawing_rels[target]
        else
          @vml_drawing_rels_id += 1
          @vml_drawing_rels[target] = @vml_drawing_rels_id
        end
      end

      def set_external_vml_links(vml_drawing_id) # :nodoc:
        @external_vml_links <<
          ['/vmlDrawing', "../drawings/vmlDrawing#{vml_drawing_id}.vml"]
      end

      def set_external_comment_links(comment_id) # :nodoc:
        @external_comment_links <<
          ['/comments',   "../comments#{comment_id}.xml"]
      end

      def external_links
        [
          @external_hyper_links,
          @external_drawing_links,
          @external_vml_links,
          @external_background_links,
          @external_table_links,
          @external_comment_links
        ].reject(&:empty?)
      end

      def drawing_links
        [@drawing_links]
      end
    end
  end
end
