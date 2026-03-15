# -*- coding: utf-8 -*-
# frozen_string_literal: true

module Writexlsx
  class Format
    class AlignmentStyle
      def initialize(format)
        @format = format
      end

      def horizontal
        @format.state.alignment.horizontal
      end

      def horizontal=(value)
        @format.state.alignment.horizontal = value
        @format.send(:sync_alignment_ivars_from_state)
      end

      def wrap
        @format.state.alignment.wrap
      end

      def wrap=(value)
        @format.state.alignment.wrap = value
        @format.send(:sync_alignment_ivars_from_state)
      end

      def vertical
        @format.state.alignment.vertical
      end

      def vertical=(value)
        @format.state.alignment.vertical = value
        @format.send(:sync_alignment_ivars_from_state)
      end

      def justlast
        @format.state.alignment.justlast
      end

      def justlast=(value)
        @format.state.alignment.justlast = value
        @format.send(:sync_alignment_ivars_from_state)
      end

      def rotation
        @format.state.alignment.rotation
      end

      def rotation=(value)
        @format.state.alignment.rotation = value
        @format.send(:sync_alignment_ivars_from_state)
      end

      def indent
        @format.state.alignment.indent
      end

      def indent=(value)
        @format.state.alignment.indent = value
        @format.send(:sync_alignment_ivars_from_state)
      end

      def shrink
        @format.state.alignment.shrink
      end

      def shrink=(value)
        @format.state.alignment.shrink = value
        @format.send(:sync_alignment_ivars_from_state)
      end

      def merge_range
        @format.state.alignment.merge_range
      end

      def merge_range=(value)
        @format.state.alignment.merge_range = value
        @format.send(:sync_alignment_ivars_from_state)
      end

      def reading_order
        @format.state.alignment.reading_order
      end

      def reading_order=(value)
        @format.state.alignment.reading_order = value
        @format.send(:sync_alignment_ivars_from_state)
      end

      def just_distrib
        @format.state.alignment.just_distrib
      end

      def just_distrib=(value)
        @format.state.alignment.just_distrib = value
        @format.send(:sync_alignment_ivars_from_state)
      end
    end
  end
end
