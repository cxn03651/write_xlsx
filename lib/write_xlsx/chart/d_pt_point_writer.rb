module Writexlsx
  class Chart
    module DPtPointWriter
      def write_d_pt_point(index, point)
        @writer.tag_elements('c:dPt') do
          write_idx(index)
          @writer.tag_elements('c:marker') do
            write_sp_pr(point)
          end
        end
      end
    end
  end
end
