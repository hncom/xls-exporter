# frozen_string_literal: true

module XlsExporter::Styler
  def default_style(**options)
    @format = Spreadsheet::Format.new(**options)
  end

  def line_height
    font_size + 3
  end

  def words_in_line
    @words_in_line || 5
  end

  def words_in_line=(count)
    @words_in_line = count
  end

  def font_size
    @font_size || 10
  end

  def font_size=(points)
    @font_size = points
  end

  def autofit(worksheet)
    fit_rows worksheet
    fit_columns worksheet
  end

  def fit_rows(worksheet)
    worksheet.rows.each do |row|
      row.height = row.each_with_index.map do |cell|
        cell.present? ? cell_height(cell) : line_height
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
    lines = words(cell).count / words_in_line
    lines += 1 if words(cell).count % words_in_line != 0
    lines * line_height
  end

  def cell_width(cell)
    words(cell).each_slice(words_in_line).map do |line|
      line.join(' ').size
    end.max
  end

  def words(string)
    string.split(' ')
  end
end
