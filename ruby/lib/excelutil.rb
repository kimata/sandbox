#!/usr/bin/env ruby
# -*- coding: utf-8 -*-

require 'win32ole'

class ExcelUtil
    def initialize
      @excel_app = WIN32OLE.new("Excel.Application")
      @excel_app.Visible = false
      @excel_app.DisplayAlerts = false
    end

    def quit
        if @excel_app then
          @excel_app.Quit
          @excel_app = nil
        end
    end

    def open(workbook_path=nil)
      workbook = nil
      if (workbook_path != nil) then
        workbook = @excel_app.Workbooks.Open({
                                               'filename' => workbook_path,
                                             })
      else
        workbook = @excel_app.Workbooks.Add
      end
      return Workbook.new(workbook)
    end

    class Workbook
      def initialize(workbook)
        @workbook = workbook
      end
      def save_as(workbook_path)
        @workbook.SaveAs({'Filename' => workbook_path})
      end
      def close(is_save_changes=false)
        @workbook.close('SaveChanges' => is_save_changes)
      end

      def get_sheet(num)
        return @workbook.Sheets.Item(num)
      end
    end
end
