namespace :data_norm do
  # rake data_norm:compare_excel['FILE_PATH_1', 'FILE_PATH_2'] --trace
  desc 'To compare excel of different years'
  task :compare_excel, [:old_file_path, :new_file_path] => :environment do |t, args|
    old_file_path = args.old_file_path
    new_file_path = args.new_file_path
    fp = FileParsing.new(old_file_path, new_file_path)
    fp.parse
  end
end
