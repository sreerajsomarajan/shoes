#
# Module Export
#
# @author [Sreeraj S]
#
# FileParsing module for excel exporting.
class FileParsing
  EXPORT_LIMIT = 1_000
  ROW_LIMIT = 64_000
  NIL_CONTENT = nil

  extend ActionView::Helpers::NumberHelper

  def initialize(old_file_path, new_file_path)
    @old_file_path = old_file_path
    @new_file_path = new_file_path
    @file_name = new_file_path.split('/').last
    parse
  end

  class << self
    def parse
      read_excel_files
      create_normalized_xlsx
    end

    def read_excel_files
      # Old excel file parsing
      @old_xlsx = Roo::Spreadsheet.open(old_file_path, extension: :xlsx)
      @old_sheet = @old_xlsx.sheet(0)
      @old_uid_array = old_sheet.column(2)
      old_col_name = @old_uid_array.shift
      return unless old_col_name == 'RPuID'

      # New excel file parsing
      @new_xlsx = Roo::Spreadsheet.open(new_file_path, extension: :xlsx)
      @new_sheet = new_xlsx.sheet(0)
      @new_uid_array = @new_sheet.column(1)
      new_col_name = @new_uid_array.shift
      return unless new_col_name == 'Previous Unique ID'
    end

    # Method to create corresponding files.
    def create_excel_file
      @file_name = new_file_path.split('/').last
      time = Time.now.strftime('%d%m%Y%M%S')
      file_name = "#{Rails.root}/public/results/#{time}.xlsx"
      FileUtils.mkdir_p(File.dirname(file_name))
      file_name
    end

    def create_normalized_xlsx(data)
      file_name = create_excel_file
      @workbook = WriteXLSX.new(file_name, font: 'Arial', size: 10)
      @worksheet = @workbook.add_worksheet('Page 1')
      write_xls_for_compare(data)
      @workbook.close
      FileUtils.rm_rf(@tmp_dir)
      file_name
    end

    def write_xls_for_compare(my_data)
      c_time = Time.now.strftime('%Y%m%d%H%M%S')
      @tmp_dir = Rails.root.join('public/temp_graph_compare/' + c_time)
      row_col = 0
      my_data.each_with_index do |ob, _i|
        merge_cells_and_write(@workbook, @worksheet, row_col, 0, row_col, 1, '')
        @worksheet.set_row(row_col, 200)
        @worksheet.set_column(row_col, 1, 25)
        unless ob[:img].include?('no_image.jpg')
          decoded_img = URI.decode(ob[:img])
          image_url = Rails.root.join('public' + decoded_img)
          image_extension = decoded_img.split('.').last
          FileUtils.mkdir_p(@tmp_dir)
          if File.exist?(image_url)
            img_name = ob[:name].tr('/', '_')
            tmp_img_url = "#{@tmp_dir}/#{img_name}.#{image_extension}"
            FastImage.resize(image_url, 360, 266, outfile: tmp_img_url)
            @worksheet.insert_image(row_col, 0, tmp_img_url, 0, 0)
          end
        end
        merge_cells_and_write(
          @workbook, @worksheet, row_col + 1, 0,
          row_col + 1, 1, ob[:name],
          color: '#FFFFFF', bg_color: '#3D3D3D'
        )
        write_xls_for_compare_detail(ob[:values], row_col)
        row_col += 23
      end
    end

    def write_xls_for_compare_detail(my_key_values, row_col)
      format_header = @workbook.add_format(
        color: '#FFFFFF', bg_color: '#acacac'
      )
      format_key = @workbook.add_format(color: 'black')
      format_value = @workbook.add_format(color: '#d2232a')
      @worksheet.write_string(row_col + 2, 0, 'ITEM', format_header)
      @worksheet.write_string(row_col + 2, 1, 'VALUE', format_header)
      row = row_col + 3
      my_key_values.each do |ob|
        @worksheet.set_row(row, 16)
        @worksheet.write_string(row, 0, ob[:key], format_key)
        @worksheet.write_string(row, 1, ob[:value], format_value)
        row += 1
      end
    end

    # Method to setup a worksheet
    def create_work_sheet(headings)
      header_format = @workbook.add_format(
        pattern: 1, size: 10, border: 1,
        color: 'black', align: 'center',
        valign: 'vcenter'
      )
      header_color = @workbook.set_custom_color(10, 171, 208, 245) # Light Blue
      header_format.set_bg_color(header_color)
      @worksheet.set_column(1, 1, 5)
      (1..6).each do |c|
        @worksheet.set_column(c, c, 20)
      end
      @worksheet.set_row(0, 20)
      xls_header_fields = headings
      @worksheet.set_column(0, 0,  20)
      @worksheet.set_column(2, 3,  20)
      @worksheet.set_column(4, 5,  15)
      @worksheet.set_column(6, 9,  10)
      xls_header_fields.each_with_index do |rec, col|
        @worksheet.write(0, col, rec, header_format)
      end
    end

    # Method to create extra excel sheet on ROW_LIMIT condition.
    def create_sheet(future_years, modifier_names)
      if @row_count % ROW_LIMIT == 0
        @i = 1
        page_number = (@row_count / ROW_LIMIT) + 1
        @worksheet = @workbook.add_worksheet("Page#{page_number}")
        headings = [
          'NAME', 'MSN', 'BUILD MONTH/YEAR',
          'ENGINE TYPE', 'MTOW (LBS)'
        ]
        headings += modifier_names.to_a
        headings += %w(CMV CBV)
        headings += future_years.to_a
        create_work_sheet(headings)
      end
    end

    def format_number(value)
      if value.to_s.include?('.')
        format_precision(format_decimal(value))
      elsif value.to_i > 1000
        format_precision(value)
      else
        value.to_s
      end
    rescue
      value.to_s
    end

    def format_decimal(value)
      format('%.2f', value.to_s)
    end

    def format_precision(value)
      number_with_delimiter(value.to_s)
    end
  end
end
