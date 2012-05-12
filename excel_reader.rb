#-*- coding: utf-8 -*-
require 'win32ole'
require 'forwardable'
require File.expand_path('../string', __FILE__)
=begin
 ExcelBookをありのままにRubyオブジェクトとして利用することを目的としている点が、ExcelModelモジュールとは異なる。
 ■利用例(読込ませたExcelの全シートのA1カラムの値を画面に出力する)
   book = ExcelReader::Book.new('c:/tmp/hoge.xls', 'A1:Z256')  # ExcelReader::Book.new(読込ませるExcelファイル, 取り込むデータの範囲[全シートで共通のため、もっとも範囲の広いシートに合わせて指定すること])
   book.sheets.each do |sheet|
     puts sheet.cell(1, 1)
   end
=end
module ExcelReader
  
  class Book
    attr_reader :sheets
    
    def initialize(options={})
      ops = {file_path: nil, range: nil, sentinel: nil}.merge(options)
      excel = WIN32OLE::new("excel.Application")
      excel.Visible = false
      
      @sheets = []
      begin
        workbook = excel.Workbooks.Open(File.expand_path(ops[:file_path].encode(Encoding::Windows_31J)))
        workbook.Sheets.each do |_sheet|
          worksheet = _sheet.respond_to?(:ole_methods) ? _sheet : workbook.Sheets.Item(_sheet["sheet_name"])
          
          @sheets << Sheet.new(worksheet, ops)
        end
      rescue
        $stderr.puts $!
      ensure
        excel.Quit
      end
    end
    
    def sheet(name)
      @sheets.select{|_sheet| _sheet.name == name}.first
    end
  end
  
  class Sheet
    extend Forwardable
    
    attr_reader :name
    
    def initialize(worksheet, options)
      @name    = worksheet.name
      @cells = Cells.new(worksheet, options)
    end
    
    def_delegator :@cells, :find, :cell
    def_delegator :@cells, :find_all, :cells
    
    def_delegators :@cells, :row_count, :column_count
  end
  
  class Cells
    attr_reader :row_count, :column_count
    
    def initialize(worksheet, options)
      @cells = {}
      worksheet.Select
      #@row_count = worksheet.Range(options[:range]).Rows.Count
      @column_count = worksheet.Range(options[:range]).Columns.Count
      worksheet.Range(options[:range]).each do |record|
        break if record.Value.to_s.encode(Encoding::UTF_8) == options[:sentinel]
        @cells.store("#{record.Row}:#{record.Column}", record.Value.to_s.encode(Encoding::UTF_8).sanitize) #Excelオブジェクトが返却する文字列はWindows-31Jなので、UTF-8に変換する。
        @row_count = record.Row #Sentinel(もしくはRangeの最終レコード)の手前のレコード位置を設定する.
      end
    end
    
    def find(row, column)
      @cells["#{row}:#{column}"]
    end
    
    def find_all
      @cells.values
    end
  end
end