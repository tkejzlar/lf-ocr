require 'rubygems'
require 'rtesseract'
require 'rmagick'
require 'string/similarity'
require 'rubyXL'
require 'rubyXL/convenience_methods'
require 'ruby/openai'
require 'base64'
require 'csv'
require 'slop'
require 'yaml'
require 'date'

# Load config
@config = YAML.load_file('config/config.yml')

# List of image paths
alliance = Dir["alliance/*.png"]
ad = Dir["ad/*.png"]
cities = Dir["cities/*.png"]
fountain = Dir["dimensions/fountain/*.png"]

@ids = {}

opts = Slop.parse do |o|
  o.bool '-ad', '--alliance-duel', 'parse alliance duel data'
  o.bool '-cp', '--combat-powers', 'parse combat power data'
  o.bool '-ct', '--cities', 'parse city capture data'
  o.bool '-df', '--dimension-fountain', 'parse Fountain of Life weekly leaderboard'
  o.string '-xls', '--excel-input', 'input excel file, unless -out is provided data will be added to it'
  o.string '-out', '--excel-output', 'output excel file, if not provided the input will be overwritten'
  o.on '--version', 'print the version' do
    puts Slop::VERSION
    exit
  end
end

opts[:excel_input] ? xls = opts[:excel_input] : xls = 'alliance.xlsx'
@workbook = RubyXL::Parser.parse(xls)
@workbook.calc_pr.full_calc_on_load = true


# Use chatGPT (4-omini) and its vision component to extract info from the cropped screenshots
def extract_info_gpt(prompt, images = [])
  gpt = OpenAI::Client.new(access_token: @config['gpt']['key'])
  messages = []
  images.each do |image|
    messages << {
      type: "image_url", 
      image_url: {
        url: "data:image/jpeg;base64,#{Base64.encode64(File.open(image.first, "rb").read)}}"
      }
    } 
  end
  messages << { type: "text", text: prompt.gsub("IDS", @ids.values.join(', ')) }
  response = gpt.chat(
        parameters: {
          model: "gpt-4o-mini",
          messages: [ { role: "user", content: messages } ],
          max_tokens: 300
        }
      )
  model_name = response.dig("model")
  prompt_tokens = response.dig("usage", "prompt_tokens")
  completion_tokens = response.dig("usage", "completion_tokens")
  total_tokens = response.dig("usage", "total_tokens")
  response_text = response.dig("choices", 0, "message", "content")
  sleep 5
  data = CSV.new(response_text).read
  if images.first.length > 1
    data = data.map { |e| e.concat(images.first.drop(1)).flatten }
  end
  return data
end

# Crop images to extract only the parts containing relevant info
def extract_player_info(image_path, index, type)
  image = Magick::Image::read(image_path).first
  index == 0 && @config['crops'].has_key?("#{type}_first") ? crop = @config['crops']["#{type}_first"] : crop = @config['crops'][type]
  crop = @config['crops'][type]
  cropped_name = "temp-cropped-#{index}.jpg"
  cropped_image = image.crop(crop[0], crop[1], crop[2], crop[3], true)
  cropped_image.write(cropped_name)
  return cropped_name
end

# Get city name - does not use GPT but manual extraction via Tesseract
def get_city_name(image_path)
  image = Magick::Image::read(image_path).first
  crop = @config['crops']['cities_name']
  cropped_image = image.crop(crop[0], crop[1], crop[2], crop[3], true);
  cropped_image.write("temp-cropped-cityname.jpg")
  ocr = RTesseract.new("temp-cropped-cityname.jpg")
  match = ocr.to_s.match(/Lv.([0-9])\s*(\S+)/)
  match == nil ? r = nil : r = [match[1].to_i, match[2]]
  return r
end

# Get current alliance members and prepare a string similarity matrix to avoid unnecessary bloating of the Excel
def populate_player_list
  players = @workbook['IDs']
  for i in 1..90 do
    id = players[i][0].value
    name = players[i][1].value
    @ids[id] = name
  end
end

# find first empty row in a spreadsheet - the workbook.sheet_data.size does not work so doing it manually by traversing the sheet and looking for first empty row
def find_first_empty_row(worksheet)
  sheet = @workbook[worksheet]
  begin
    for i in 1..100000 do
      val = sheet[i][0].value
      return i if val.empty?
    end
  rescue
    return i
  end
end

# recheck each player name against existing names performing cosine similarity, replace when > 90% match
def sanitize_player_name(str)
  clean_name = str.gsub(/\W*/,'')
  similarities = {}
  ids = @ids.values
  ids.each_with_index do |name, i|
    sim = String::Similarity.cosine(clean_name, name)
    similarities[sim] = [name, i]
  end
  possible_match = similarities.sort.last
  if possible_match[0] > 0.9
    clean_name = possible_match[1][0]
  end
  return clean_name
end

# print out a tab-separated list of extracted data to be pasted into Excel
def print_data(output, *columns)
  columns.reject!(&:empty?)
  output.each do |row|
    next if row[0][0,3] == '```'
    name = sanitize_player_name(row[0])
    if columns.length > 0
      puts "#{name}\t#{row.drop(1).join("\t")}\t#{columns.join("\t")}\t#{Time.now.strftime("%d/%m/%Y")}"
    else
      puts "#{name}\t#{row.drop(1).join("\t")}\t#{Time.now.strftime("%d/%m/%Y")}"
    end
  end
end

# write data into excel
def write_data_excel(type, data, *columns)
  worksheet = @config['excel'][type]['sheet']
  sheet = @workbook[worksheet]
  first_empty = find_first_empty_row(worksheet)
  data.select{ |k,v| k[0,3] != '```' }.to_a.each_with_index do |row, i|
    rowno = first_empty + i
    s = 0
    sheet.add_cell(rowno, 0, nil, @config['excel'][type]['formula-a'].gsub('$ROW', (rowno + 1).to_s))
    sheet.add_cell(rowno, 1, sanitize_player_name(row[0]))
    sheet.add_cell(rowno, 2, row[1].to_i)
    row.drop(2).each_with_index do |r, s|
      sheet.add_cell(rowno, 2 + s, r)
    end
    columns.flatten.each_with_index do |cell, j|
      if cell[0,7] == 'formula'
        sheet.add_cell(rowno, 3+s+j, nil, @config['excel'][type][cell].gsub('$ROW', (rowno + 1).to_s))
      elsif cell == 'date'
        c = sheet.add_cell(rowno, 3+s+j)
        c.set_number_format('dd.mm.YYYY')
        c.change_contents(Date.today)
      else
        sheet.add_cell(rowno, 3+s+j, cell)
      end
    end
  end
end

# Hook for getting city names and levels
def hook_cities(image)
  r = get_city_name(image)
  return r if r != nil
end

# clear any temporary cropped images
def clear_temp
  Dir.glob(["*.jpg"]) { |f| File.delete(f) }
end

# process scores for a given type of images
def process_scores(type, crops = nil, *additional_fields)
  images = []
  output = []
  more_fields = []
  crops = type if crops == nil
  puts "===== #{crops.upcase} ====="
  raw_images = Dir["#{type}/*.png"]
  raw_images.each_with_index do |path, i|
    begin
      r = send("hook_#{type}", path)
      if r != nil
        more_fields = r
        next
      end
    rescue
    end
    images << [extract_player_info(path, i, crops), more_fields]
  end
  images.each_slice(3) do |slice|
    output.concat(extract_info_gpt(@config['gpt']['prompts'][type], slice))
  end
  print_data(output, additional_fields)
  formulas = @config['excel'][type].select { |k, v| k.include?('formula') && k != 'formula-a'}
  write_data_excel(type, output, additional_fields, 'date', formulas.keys)
  clear_temp
end

# populate player list
populate_player_list

# Loop over all images and extract information
if opts.combat_powers?
  process_scores('alliance')
end

if opts.alliance_duel?
  process_scores('ad')
end

if opts.cities?
  process_scores('cities', 'cities_rally', 'siege')
  process_scores('cities', 'cities_dm', 'dm')
end

opts[:excel_output] ? out = opts[:excel_output] : out = xls
@workbook.write(out)