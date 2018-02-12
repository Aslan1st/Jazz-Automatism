require 'win32ole'
require 'Launchy'
require 'Date'
require 'nokogiri'

class BroadcastSession

  def initialize
    update_title
    update_output
    Launchy.open('C:\Users\Aslan\Desktop\Mezzrow_Live_Stream.xml')
  end

  def bridge
    new_wirecast = WIN32OLE.connect('Wirecast.Application')
    @wirecast_documentobject = new_wirecast.DocumentByName('Mezzrow_Live_Stream', 0)
    @document_layer =  @wirecast_documentobject.LayerByName('normal')
  end

  def start_stream
    if @wirecast_documentobject.IsBroadcasting() == 0
      @wirecast_documentobject.Broadcast('start')
    end
  end

  def start_record_to_disk
    if @wirecast_documentobject.IsArchivingToDisk() == 0
      @wirecast_documentobject.ArchiveToDisk('start')
    end
  end

  def start_broadcast
    begin
      is_ok = true
      if is_ok == true
        bridge
      end
    rescue WIN32OLERuntimeError
      is_ok = false
    end
      if is_ok == false
        sleep(1)
        start_broadcast
      end
    @shot_index = 1
    @total_shots = @document_layer.ShotCount()
  end

  def stop_broadcast
    @wirecast_documentobject.Broadcast('stop')
    @wirecast_documentobject.ArchiveToDisk('stop')
  end

  def update_xml
    update_title
    update_output
  end

  def update_title
    file_name = 'C:\Users\Aslan\Desktop\Mezzrow_Live_Stream.xml'

    doc = Nokogiri::XML(File.open(file_name))
    shot = doc.at('shot[unique_id="f3341230-4a49-40f8-a0c7-e7784e6103ef_shot"]')
    input = shot.at('text').attribute('text')
    input.content = DateTime.now.strftime('Mezzrow ' + '%Y' + '-' + '%m' + '-' + '%d' + ' © 2017- smallslive.com')

    File.open(file_name, 'w') {|f| f.puts doc.to_xml}
  end

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

##
#  def name_files
#    folder_path = 'C:/Users/Aslan/Desktop/Automatism'
#    todays_date =  Date.today.strftime('%Y-%m-%d')
#    Dir[folder_path + '/*'].reject{ |f| f['C:/Users/Aslan/Desktop/Automatism/TestDir'] }.each do |f|
#      filename = File.basename(f, File.extname(f)).tr('_0', '')
#      filename_new = [ folder_path,
#                       '/',
#                       'TestDir/',
#                       filename,
#                       todays_date,
#                       File.extname(f) ].map!(&:to_s).join("")
#      File.rename(f, filename_new)
#    end
#  end

  def rotate_shots(new_shot)
    @document_layer.setproperty('ActiveShotID', @document_layer.ShotIDByIndex(new_shot))
  end

  def which_shot
    if @shot_index == @total_shots
      @shot_index = 1
    end
    @shot_index += 1
  end

  def get_ending_datetime
    calculate_tomorow = Date.today + 1
    date_tomorrow = calculate_tomorow.strftime('%Y-%m-%d') + 'T04:00:00-05:00'
    @four_am_tomorrow = DateTime.xmlschema(date_tomorrow)
  end

  def calculate_switch
    @when_to_change = Time.now + 20
  end

  def good_to_change
    return @when_to_change <= Time.now
  end

  def get_up_and_go

      start_broadcast
      #start_stream
      start_record_to_disk
      get_ending_datetime
      while DateTime.now <= @four_am_tomorrow do
        calculate_switch
        rotate_shots(which_shot)
        until good_to_change
          sleep(1)
        end
      end
      stop_broadcast
  end

end
