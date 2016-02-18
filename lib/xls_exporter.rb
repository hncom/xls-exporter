require 'xls_exporter/version'
require 'spreadsheet'

module XlsExporter
  class Exporter
    def self.export(&block)
      exporter = new
      exporter.instance_exec(&block)
      exporter.save!
    end

    def initialize
      @book     = Spreadsheet::Workbook.new
      @filename = "default_filename"
    end

    def add_sheet(sheet_name = nil)
      save_sheet! if @sheet
      @headers = []
      @body    = []
      @sheet   = @book.create_worksheet

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

    def export_models(scope, *attrs)
      # TODO: associate human name with proc in attrs
      headers(*attrs.map(&:to_s).map(&:humanize))
      to_body = scope.map do |instance|
        attrs.map do |attr|
          if attr.class == Proc
            attr.call instance
          elsif attr.class == Symbol
            # shouldn't using `try` because of SimpleDelegator
            instance.respond_to?(attr) ? instance.send(attr) : nil
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
      filename = "./#{@filename}_#{Time.now.to_i}.xls"
      @book.write(filename)
      puts "Report has been saved as #{filename}"
    end
  end
end
