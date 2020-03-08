# frozen_string_literal: true

require 'spreadsheet'
require 'xls_exporter/styler'

class XlsExporter::Exporter
  include XlsExporter::Styler

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

  def filename(new_filename)
    @filename = new_filename
  end

  def headers(*args)
    @headers = args
  end

  def body(new_body)
    @body = new_body
  end

  def humanize_columns(columns)
    columns.map do |column|
      column = column.keys.first if column.is_a? Hash
      column.to_s.humanize
    end
  end

  def export_models(scope, *columns)
    headers(*humanize_columns(columns))
    to_body = scope.map do |instance|
      columns.map do |column|
        column = column.values.first if column.is_a? Hash
        if column.is_a? Proc
          instance.instance_exec(&column)
        elsif column.is_a? Symbol
          instance.send column
        end
      end
    end
    body to_body
  end

  def save_sheet!
    @sheet.row(0).concat(@headers)
    @body.each_with_index do |row, index|
      @sheet.row(index + 1).concat(row)
    end
  end

  def save!
    save_sheet!
    @book.worksheets.each do |worksheet|
      autofit worksheet
    end
    if @filename.present?
      filename = "./#{@filename}_#{Time.now.to_i}.xls"
      @book.write(filename)
      puts "Report has been saved as #{filename}"
    else
      @book
    end
  end

  WORDS_IN_LINE = 5
  FONT_SIZE = 10
  LINE_HEIGHT = FONT_SIZE + 3

  def autofit(worksheet)
    worksheet.rows.each do |row|
      lines_count = row.each_with_index.map do |cell, _index|
        if cell.present?
          words_count = cell.split(' ').count
          lines = words_count / WORDS_IN_LINE
          lines += 1 if words_count % WORDS_IN_LINE != 0
          lines * LINE_HEIGHT
        else
          LINE_HEIGHT
        end
      end
      row.height = lines_count.max
    end
    worksheet.column_count.times do |col_idx|
      column = worksheet.column(col_idx)
      column.width = column.each_with_index.map do |cell, _row|
        if cell.present?
          words = cell.split(' ')
          words.each_slice(WORDS_IN_LINE).map do |line|
            line.join(' ').size
          end.max
        else
          0
        end
      end.max
    end
  end
end
