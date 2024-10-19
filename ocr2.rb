require 'rubygems'
require 'rtesseract'
require 'rmagick'
require 'string/similarity'
require 'rubyXL'
require 'ruby/openai'
require 'base64'
require 'csv'
require 'slop'
require 'yaml'

# Load config
@config = YAML.load_file('config/config.yml')

# List of image paths and crops
alliance = Dir["alliance/*.png"]
ad = Dir["ad/*.png"]
cities = Dir["cities/*.png"]
@crops = {
  'alliance' => [666,116,(2556-666),(1017-116)], 
  'alliance_first' => [666,416,(2556-666),(1017-416)],
  'ad' => [738,397,(2108-738),(888-397)], 
  'cities_rally' => [844, 148, (1557-844), (990-148)], 
  'cities_dm' => [1681, 148, (2270-1681), (990-148)]
}

@ids = {}

@workbook = RubyXL::Parser.parse("alliance.xlsx")
sheet = @workbook['CPs']
for i in 0..100000 do
  cell = sheet[i][1].value
 
  return i if cell.empty?
end
Kernel.exit

opts = Slop.parse do |o|
  o.bool '-ad', '--alliance-duel', 'parse alliance duel data'
  o.bool '-cp', '--combat-powers', 'parse combat power data'
  o.bool '-ct', '--cities', 'parse city capture data'
  o.on '--version', 'print the version' do
    puts Slop::VERSION
    exit
  end
end

# Crop images to extract only the parts containing relevant info
def extract_player_info(image_path, index, type)
  image = Magick::Image::read(image_path).first
  index == 0 && @crops.has_key?("#{type}_first") ? crop = @crops["#{type}_first"] : crop = @crops[type]
  crop = @crops[type]
  cropped_name = "temp-cropped-#{index}.jpg"
  cropped_image = image.crop(crop[0], crop[1], crop[2], crop[3], true);
  cropped_image.write(cropped_name)
  return cropped_name
end

# Use chatGPT (4-omini) and its vision component to extract info from the cropped screenshots
def extract_info_gpt(prompt, images = [])
  gpt = OpenAI::Client.new(access_token: @config['gpt'])
  
  messages = []
  images.each do |image|
    messages << {
      type: "image_url", 
      image_url: {
        url: "data:image/jpeg;base64,#{Base64.encode64(File.open(image, "rb").read)}}"
      }
    } 
  end
  
  messages << { type: "text", text: prompt }

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
  return CSV.new(response_text).read
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

# recheck each player name against existing names performing cosine similarity, replace when > 70% match
def sanitize_player_name(str)
  clean_name = str.gsub(/\W*/,'')
  similarities = {}
  ids = @ids.values
  ids.each_with_index do |name, i|
    sim = String::Similarity.cosine(clean_name, name)
    similarities[sim] = [name, i]
  end
  possible_match = similarities.sort.last
  if possible_match[0] > 0.7
    clean_name = possible_match[1][0]
  end
  return clean_name
end

# print out a tab-separated list of extracted data to be pasted into Excel
def print_data(output, *columns)
  columns.reject!(&:empty?)
  output.each do |name, score|
    next if name[0,3] == '```'
    name = sanitize_player_name(name)
    if columns.length > 0
      puts "#{name}\t#{score}\t#{columns.join("\t")}\t#{Time.now.strftime("%d/%m/%Y")}"
    else
      puts "#{name}\t#{score}\t#{Time.now.strftime("%d/%m/%Y")}"
    end
  end
end

# remove any temporary cropped images
def clear_temp
  Dir.glob(["*.jpg"]) { |f| File.delete(f) }
end

# find first empty row in a spreadsheet - the workbook.sheet_data.size does not work so doing it manually by traversing the sheet and looking for first empty row
def find_empty_row(worksheet)
  sheet = @workbook[worksheet]
  for i in 0..100000 do
    cell = sheet[i][0].value
    return i if cell.value.empty?
  end
end

def add_excel_row(output, worksheet, first_empty, *columns)
  sheet = @workbook[worksheet]
  formula = sheet[first_empty-1][0].formula
  output.to_a.each_with_index do |row, i|
    sheet[first_empty+i][0].change_contents("", formula)
    sheet[first_empty+i][1].change_contents(row[0])
    sheet[first_empty+i][2].change_contents(row[1])
    columns.each_with_index do |cell, j|
      sheet[first_empty+i][3+j].change_contents(cell)
    end
  end
end

# populate player list
populate_player_list

# prompts used for chatGPT
@prompts = {
  'alliance' => "The images I shared contain screenshots from the game Last Fortress. They represent a list of players in an alliance and their combat power (CP). All numeric values are using comma as a thousands separator. Please extract all players and their CPs (full scores, not rounded to millions or anything) in a CSV-compatible format. Your response should ONLY contain the comma separated list of player names and CPs. Also, for player names, here is the actual player list - please use it to correct what you read from the images to get consistent names: #{@ids.values.join(', ')}",
  'ad' => "The images I shared contain screenshots from the game Last Fortress. They represent a list of players in an alliance and their contributions to alliance duel. All numeric values are using comma as a thousands separator. Please extract all players and their alliance duel scores (full scores, not rounded to millions or anything) in a CSV-compatible format. Your response should ONLY contain the comma separated list of player names and CPs. Also, for player names, here is the actual player list - please use it to correct what you read from the images to get consistent names: #{@ids.values.join(', ')}", 
  'cities' => "The images I shared contain screenshots from the game Last Fortress. They represent a list of players that contributed to attacking a city. There always is player name, contribution score, and for the top 3 players also merits. All numeric values are using comma as a thousands separator. Please extract all players and their contribution scores (full scores, not rounded to millions or anything, also ignore the merit scores if they exist) in a CSV-compatible format. Your response should ONLY contain the comma separated list of player names and contribution scores. Also, for player names, here is the actual player list - please use it to correct what you read from the images to get consistent names: #{@ids.values.join(', ')}"
}

# Loop over all images and extract information
# TODO: this can be probably refactored to avoid duplication
if opts.combat_powers?
  puts "===== PLAYER CPs ====="
  images = []
  output = []
  alliance.each_with_index do |path, i|
    images << extract_player_info(path, i, 'alliance')
  end
  images.each_slice(3) do |slice|
    output.concat(extract_info_gpt(@prompts['alliance'], slice))
  end
  print_data(output)
  clear_temp
end

if opts.alliance_duel?
  puts "===== AD SCORES ====="
  images = []
  output = []
  ad.each_with_index do |path, i|
    images << extract_player_info(path, i, 'ad')
  end
  images.each_slice(3) do |slice|
    output.concat(extract_info_gpt(@prompts['ad'], slice))
  end
  print_data(output)
  clear_temp
end

if opts.cities?
  puts "===== CITY RALLY SCORES ====="
  images = []
  output = []
  cities.each_with_index do |path, i|
    images << extract_player_info(path, i, 'cities_rally')
  end
  images.each_slice(3) do |slice|
    output.concat(extract_info_gpt(@prompts['cities'], slice))
  end
  #print_data(output, 'siege')
  add_excel_row(output, 'Cities', find_empty_row('Cities'), 'siege')
  clear_temp

  puts "===== CITY DM SCORES ====="
  images = []
  output = []
  cities.each_with_index do |path, i|
    images << extract_player_info(path, i, 'cities_dm')
  end
  images.each_slice(3) do |slice|
    output.concat(extract_info_gpt(@prompts['cities'], slice))
  end
  print_data(output, 'dm')
  clear_temp
end
