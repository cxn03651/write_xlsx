# -*- coding: utf-8 -*-
# frozen_string_literal: true

module Writexlsx
  class Format
    class AlignmentStyle
      def initialize(format)
        @format = format
      end

      def horizontal
        @format.instance_variable_get(:@alignment_state).horizontal
      end

      def horizontal=(value)
        @format.instance_variable_get(:@alignment_state).horizontal = value
        @format.send(:sync_alignment_ivars_from_state)
      end

      def wrap
        @format.instance_variable_get(:@alignment_state).wrap
      end

      def wrap=(value)
        @format.instance_variable_get(:@alignment_state).wrap = value
        @format.send(:sync_alignment_ivars_from_state)
      end

      def vertical
        @format.instance_variable_get(:@alignment_state).vertical
      end

      def vertical=(value)
        @format.instance_variable_get(:@alignment_state).vertical = value
        @format.send(:sync_alignment_ivars_from_state)
      end

      def justlast
        @format.instance_variable_get(:@alignment_state).justlast
      end

      def justlast=(value)
        @format.instance_variable_get(:@alignment_state).justlast = value
        @format.send(:sync_alignment_ivars_from_state)
      end

      def rotation
        @format.instance_variable_get(:@alignment_state).rotation
      end

      def rotation=(value)
        @format.instance_variable_get(:@alignment_state).rotation = value
        @format.send(:sync_alignment_ivars_from_state)
      end

      def indent
        @format.instance_variable_get(:@alignment_state).indent
      end

      def indent=(value)
        @format.instance_variable_get(:@alignment_state).indent = value
        @format.send(:sync_alignment_ivars_from_state)
      end

      def shrink
        @format.instance_variable_get(:@alignment_state).shrink
      end

      def shrink=(value)
        @format.instance_variable_get(:@alignment_state).shrink = value
        @format.send(:sync_alignment_ivars_from_state)
      end

      def merge_range
        @format.instance_variable_get(:@alignment_state).merge_range
      end

      def merge_range=(value)
        @format.instance_variable_get(:@alignment_state).merge_range = value
        @format.send(:sync_alignment_ivars_from_state)
      end

      def reading_order
        @format.instance_variable_get(:@alignment_state).reading_order
      end

      def reading_order=(value)
        @format.instance_variable_get(:@alignment_state).reading_order = value
        @format.send(:sync_alignment_ivars_from_state)
      end

      def just_distrib
        @format.instance_variable_get(:@alignment_state).just_distrib
      end

      def just_distrib=(value)
        @format.instance_variable_get(:@alignment_state).just_distrib = value
        @format.send(:sync_alignment_ivars_from_state)
      end
    end
  end
end
