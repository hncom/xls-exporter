# frozen_string_literal: true

module XlsExporter::Styler
  def set_default_style(**options)
    @format = Spreadsheet::Format.new(**options)
  end
end
