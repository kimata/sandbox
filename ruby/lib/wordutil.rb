#!/usr/bin/env ruby
# -*- coding: utf-8 -*-

Encoding.default_internal = 'UTF-8'

require 'win32ole'
require 'word/const'

module MSO
end

WIN32OLE.const_load('Microsoft Office 15.0 Object Library', MSO)

class WordUtil
    def initialize
      @word_app = WIN32OLE.new("Word.Application")
      @word_app.visible = false
    end

    def quit
        if @word_app then
          @word_app.Quit(false)
          @word_app = nil
        end
    end

    def open(doc_path)
      full_path = File.expand_path(doc_path)
      doc = @word_app.Documents.Open({
                                       'filename' => full_path,
                                       'readOnly' => true
                                     })
      return Document.new(doc)
    end

    class Document
      def initialize(doc)
        @doc = doc
      end

      def get_total_page_number
        return @doc.ComputeStatistics(Word::Const::WdStatisticPages)
      end
    end
end


#     def open


#     def parse(file_path)
#       full_path = File.expand_path(file_path)
#       # doc = @word_app.Documents.Open(full_path)

# # MSO.constants.each {|v|
# #   print '    ', v, ' = ', eval("MSO::#{v}"), "\n"
# # }

#       doc = @word_app.Documents.Open({
#                                        'filename' => full_path,
#                                        'readOnly' => true
#                                      })

#       # doc.Paragraphs.each {|paragraph|
#       #   text  = paragraph.Range.Text
#       #   puts text
#       # }


#       doc.Shapes.each {|shape|
#         next if (shape.Type != MSO::MsoTextBox) 
#         next if (shape.TextFrame.TextRange.Style.NameLocal != SSR_STYLE)



#         puts shape.TextFrame.TextRange.Text

#         puts '---------'

#         puts shape.Anchor.Information(WORD::CONST::WdActiveEndPageNumber);

#       }
#       title = ''
#       id = ''
#       doc.BuiltInDocumentProperties.each {|prop|
#         title = prop.Value if (prop.Name == 'Title') 
#         id = prop.Value if (prop.Name == 'Subject') 
#       }

#         puts title
#         puts id
#         puts 'END'
#     end
# end

# doc_path = ARGV[0]
# puts "input doc: " + doc_path

  
# begin
#   parser = WordParser.new()
#   parser.parse(doc_path)
# ensure
#   parser.quit  
# end


















# # convert to pdf
# WordManager.run_during do |word|
#     Dir.glob(input_docs) do |fn|
#         word.to_pdf(fn)
#     end
# end


