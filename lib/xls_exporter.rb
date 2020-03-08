# frozen_string_literal: true

require 'xls_exporter/version'
require 'xls_exporter/exporter'

module XlsExporter
  def self.export(&block)
    Exporter.export(&block)
  end
end
