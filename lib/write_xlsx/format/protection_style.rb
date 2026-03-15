# -*- coding: utf-8 -*-
# frozen_string_literal: true

module Writexlsx
  class Format
    class ProtectionStyle
      def initialize(format)
        @format = format
      end

      def locked
        @format.instance_variable_get(:@protection_state).locked
      end

      def locked=(value)
        @format.instance_variable_get(:@protection_state).locked = value
        @format.send(:sync_protection_ivars_from_state)
      end

      def hidden
        @format.instance_variable_get(:@protection_state).hidden
      end

      def hidden=(value)
        @format.instance_variable_get(:@protection_state).hidden = value
        @format.send(:sync_protection_ivars_from_state)
      end
    end
  end
end
