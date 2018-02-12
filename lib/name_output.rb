require 'nokogiri'

def update_output
  file_name = 'C:\Users\Aslan\Desktop\Mezzrow_Live_Stream.xml'

  doc = Nokogiri::XML(File.open(file_name))
  output_node = doc.at("advanced_output")
  hr_output = output_node.at('output[unique_id="2"]').attribute('output_url')
  lr_output = output_node.at('output[unique_id="3"]').attribute('output_url')
  hr_output.content = name_files("HR")
  lr_output.content = name_files("LR")

  File.open(file_name, 'w') {|f| f.puts doc.to_xml}
end

def name_files(quality)
  folder_path = 'C:\\Users\\Aslan\\Desktop\\Automatism'
  todays_date =  Date.today.strftime('%Y-%m-%d')
  filename_new = [ folder_path,
                    '\\',
                    'TestDir\\',
                    quality,
                    todays_date,
                    '.mp4' ].map!(&:to_s).join("")
end
