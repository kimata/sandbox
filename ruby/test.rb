#!/usr/bin/env ruby
# -*- coding: utf-8 -*-

$LOAD_PATH.push('./lib')

require 'wordutil'
require 'findutil'

wordutil = WordUtil.new

path = 'D:\\Document\\ssr_info\\sample'

FindUtil.find(path, {
                file_filter: lambda {|name| /^[^~].*\.doc$/.match(name) },
                dir_filter: lambda {|name| !/old/.match(name) },
              }) {|path, rel_path|

  doc = wordutil.open(path)
  puts "name: %-20s, page:%d\n" % [Pathname.new(path).basename, doc.get_total_page_number]
}

