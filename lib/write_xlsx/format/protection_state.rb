# -*- coding: utf-8 -*-
# frozen_string_literal: true

module Writexlsx
  class Format
    class ProtectionState
      attr_accessor :locked, :hidden

      def initialize
        @locked = 1
        @hidden = 0
      end

      def initialize_copy(other)
        @locked = other.locked
        @hidden = other.hidden
      end
    end
  end
end
