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

  attr_accessor :old_file_path, :new_file_path

  extend ActionView::Helpers::NumberHelper

  def initialize(old_file_path, new_file_path)
    @old_file_path = Rails.root.to_s + old_file_path.to_s
    @new_file_path = Rails.root.to_s + new_file_path.to_s
    @file_name = @new_file_path.split('/').last.split('.').first
  end

  def parse
    read_excel_files
    create_normalized_xlsx
  end

  def read_excel_files
    # Old excel file parsing
    @old_xlsx = Roo::Spreadsheet.open(@old_file_path, extension: :xlsx)
    @old_sheet = @old_xlsx.sheet(0)
    @old_uid_array = @old_sheet.column(2)
    old_col_name = @old_uid_array.shift
    return unless old_col_name == 'RPuID'

    # New excel file parsing
    @new_xlsx = Roo::Spreadsheet.open(@new_file_path, extension: :xlsx)
    @new_sheet = @new_xlsx.sheet(0)
    @new_uid_array = @new_sheet.column(1)
    new_col_name = @new_uid_array.shift
    return unless new_col_name == 'Previous Unique ID'
  end

  def create_normalized_xlsx
    file_name = create_excel_file
    @alive_row_count = 0
    @missing_row_count = 0
    @workbook = WriteXLSX.new(file_name, font: 'Arial', size: 10)
    @alive_sheet = @workbook.add_worksheet('Alive')
    @missing_sheet = @workbook.add_worksheet('Missing')
    core_compare
    @workbook.close
  end

  # Method to create corresponding files.
  def create_excel_file
    time = Time.now.strftime('%d%m%Y%M%S')
    file_name = "#{Rails.root}/public/results/#{@file_name}_#{time}.xlsx"
    FileUtils.mkdir_p(File.dirname(file_name))
    file_name
  end

  def core_compare
    @new_sheet.each_with_index do |row, index|
      if index.zero?
        @new_header = row
        @old_header = @old_sheet.row(index + 1)
        create_work_sheet
      else
        if @old_uid_array.include?(row[0])
          @alive_sheet.write_col(@alive_row_count, 0, [row.map(&:to_s)])
          @alive_row_count += 1
        else
          @missing_sheet.write_col(@missing_row_count, 0, [row.map(&:to_s)])
          @missing_row_count += 1
        end
      end
      # break if index == 2
    end
  end

  # Method to setup a worksheet
  def create_work_sheet
    header_format = @workbook.add_format(
      pattern: 1, size: 10, border: 1,
      color: 'black', align: 'center',
      valign: 'vcenter'
    )
    header_color = @workbook.set_custom_color(10, 171, 208, 245) # Light Blue
    header_format.set_bg_color(header_color)

    # Alive sheet
    @alive_sheet.set_column(0, @new_header.length, 20)
    @alive_sheet.set_row(0, 20)
    @alive_sheet.write_col(@alive_row_count, 0, [@new_header], header_format)
    @alive_row_count += 1

    # Missing sheet
    @missing_sheet.set_column(0, @old_header.length, 20)
    @missing_sheet.set_row(0, 20)
    @missing_sheet.write_col(@missing_row_count, 0, [@old_header], header_format)
    @missing_row_count += 1
  end
end

# require 'rubygems'
# require 'active_support'
# require 'roo'
# require 'writeexcel'
# require 'fileutils'
# require 'rails'
# # @old_file_path = ARGV[0]
# # @new_file_path = ARGV[1]
# #
# # Model FileParsing
# #
# # @author [Sreeraj S]
# #
# class FileParsing
#   def self.parse(@old_file_path, @new_file_path)
#     @file_name = @new_file_path.split('/').last
#     read_excel_files
#     create_normalized_xlsx
#   end

#   def self.create_normalized_xlsx
#     @workbook = WriteXLSX.new(@file_name, font: 'Arial', size: 10)
#     @alive_sheet = @workbook.add_worksheet('Alive')
#     @missing_sheet = @workbook.add_worksheet('Missing')
#     core_compare
#     @workbook.close
#   end

#   def read_excel_files
#     # Old excel file parsing
#     @old_xlsx = Roo::Spreadsheet.open(@old_file_path, extension: :xlsx)
#     @old_sheet = @old_xlsx.sheet(0)
#     @old_uid_array = old_sheet.column(2)
#     old_col_name = @old_uid_array.shift
#     return unless old_col_name == 'RPuID'

#     # New excel file parsing
#     @new_xlsx = Roo::Spreadsheet.open(@new_file_path, extension: :xlsx)
#     @new_sheet = new_xlsx.sheet(0)
#     @new_uid_array = @new_sheet.column(1)
#     new_col_name = @new_uid_array.shift
#     return unless new_col_name == 'Previous Unique ID'
#   end

#   def core_compare
#     @new_sheet.each_with_index do |row, index|
#       @alive_sheet.write_string(@i, column, record['amp'].to_s)
#       @alive_sheet.add_row(row) if index.zero?
#       if @old_uid_array.include?(row[0])
#         p 'ROW'
#         p row
#         @alive_sheet.add_row(row)
#       else
#         @missing_sheet.add_row(row)
#       end
#       break if index == 2
#     end
#     folder_path = (Rails.root.to_s + '/public/results/')
#     FileUtils.mkdir_p(folder_path)
#     FileUtils.mv @doc.path, folder_path + @file_name
#   end

#   # Method to setup a worksheet
#   def create_work_sheet(headings)
#     header_format = @workbook.add_format(
#       pattern: 1, size: 10, border: 1,
#       color: 'black', align: 'center',
#       valign: 'vcenter'
#     )
#     header_color = @workbook.set_custom_color(10, 171, 208, 245) # Light Blue
#     header_format.set_bg_color(header_color)
#     @worksheet.set_column(1, 1, 5)
#     (1..6).each do |c|
#       @worksheet.set_column(c, c, 20)
#     end
#     @worksheet.set_row(0, 20)
#     xls_header_fields = headings
#     @worksheet.set_column(0, 0,  20)
#     @worksheet.set_column(2, 3,  20)
#     @worksheet.set_column(4, 5,  15)
#     @worksheet.set_column(6, 9,  10)
#     xls_header_fields.each_with_index do |rec, col|
#       @worksheet.write(0, col, rec, header_format)
#     end
#   end

#   # Method to create extra excel sheet on ROW_LIMIT condition.
#   def create_sheet(future_years, modifier_names)
#     if @row_count % ROW_LIMIT == 0
#       @i = 1
#       page_number = (@row_count / ROW_LIMIT) + 1
#       @worksheet = @workbook.add_worksheet("Page#{page_number}")
#       headings = [
#         'NAME', 'MSN', 'BUILD MONTH/YEAR',
#         'ENGINE TYPE', 'MTOW (LBS)'
#       ]
#       headings += modifier_names.to_a
#       headings += %w(CMV CBV)
#       headings += future_years.to_a
#       create_work_sheet(headings)
#     end
#   end
# end
