# frozen_string_literal: true

module XlsExporter::Styler
  def default_style=(**options)
    @format = Spreadsheet::Format.new(**options)
  end

  WORDS_IN_LINE = 5
  FONT_SIZE = 10
  LINE_HEIGHT = FONT_SIZE + 3

  def autofit(worksheet)
    fit_rows worksheet
    fit_columns worksheet
  end

  def fit_rows(worksheet)
    worksheet.rows.each do |row|
      row.height = row.each_with_index.map do |cell|
        cell.present? ? cell_height(cell) : LINE_HEIGHT
      end.max
    end
  end

  def fit_columns(worksheet)
    worksheet.column_count.times do |col_idx|
      column = worksheet.column(col_idx)
      column.width = column.each_with_index.map do |cell|
        cell.present? ? cell_width(cell) : 0
      end.max
    end
  end

  def cell_height(cell)
    lines = words(cell).count / WORDS_IN_LINE
    lines += 1 if words(cell).count % WORDS_IN_LINE != 0
    lines * LINE_HEIGHT
  end

  def cell_width(cell)
    words(cell).each_slice(WORDS_IN_LINE).map do |line|
      line.join(' ').size
    end.max
  end

  def words(string)
    string.split(' ')
  end
end
