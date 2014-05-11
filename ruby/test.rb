#!/usr/bin/env ruby
# -*- coding: utf-8 -*-

$LOAD_PATH.push('./lib')

require 'wordutil'
require 'excelutil'
require 'findutil'



path = 'D:\\Document\\ssr_info\\sample'
excel_path = 'D:\\Document\\ssr_info\\sample\\info.xls'

page_info = {}

begin
  wordutil = WordUtil.new
  FindUtil.find(path, {
                  file_filter: lambda {|name| /^[^~].*\.doc$/.match(name) },
                  dir_filter: lambda {|name| !/old/.match(name) },
                }) {|path, rel_path|
    doc = wordutil.open(path)
    page_info[Pathname.new(path).basename.to_s] = doc.get_total_page_number
    doc.close
  }
ensure
  wordutil.quit  
end



excelutil = ExcelUtil.new

begin
  workbook = excelutil.open
  sheet = workbook.get_sheet(1)

  sheet.rows[1].columns[1] = 'ファイル'
  sheet.rows[1].columns[2] = 'ページ数'


  page_info.keys.each_with_index{|key, i| 
    sheet.rows[2+i].columns[1] = key
    sheet.rows[2+i].columns[2] = page_info[key]
  }
  workbook.save_as(excel_path)
ensure
  excelutil.quit  
end



