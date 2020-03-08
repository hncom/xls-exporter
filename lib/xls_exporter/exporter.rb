# frozen_string_literal: true

require 'spreadsheet'
require 'xls_exporter/styler'

class XlsExporter::Exporter
  include XlsExporter::Styler

  attr_writer :filename, :body

  def self.export(&block)
    exporter = new
    exporter.instance_exec(&block)
    exporter.save!
  end

  def initialize
    @book = Spreadsheet::Workbook.new
  end

  def add_sheet(sheet_name = nil)
    save_sheet! if @sheet
    @headers = []
    @body    = []
    @sheet   = @book.create_worksheet

    @sheet.default_format = @format if @format.present?
    @sheet.name = sheet_name
  end

  def headers(*args)
    @headers = args
  end

  def humanize_columns(columns)
    columns.map do |column|
      column = column.keys.first if column.is_a? Hash
      column.to_s.humanize
    end
  end

  def export_models(scope, *columns)
    headers(*humanize_columns(columns))
    body to_body scope, columns
  end

  def to_body(scope, columns)
    scope.map do |instance|
      columns.map do |column|
        column = column.values.first if column.is_a? Hash
        if column.is_a? Proc
          instance.instance_exec(&column)
        elsif column.is_a? Symbol
          instance.send column
        end
      end
    end
  end

  def save_sheet!
    @sheet.row(0).concat(@headers)
    @body.each_with_index do |row, index|
      @sheet.row(index + 1).concat(row)
    end
  end

  def save!
    save_sheet!
    @book.worksheets.each { |worksheet| autofit worksheet }
    if @filename.present?
      filename = "./#{@filename}_#{Time.now.to_i}.xls"
      @book.write(filename)
      puts "Report has been saved as #{filename}"
    else
      @book
    end
  end
end
