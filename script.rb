require "HTTParty"
require "Nokogiri"
require "JSON"
require "pry"
require "google_drive"
require "xmlhasher"
require 'active_support/core_ext/hash/conversions'
require 'mechanize'
require "yaml"

config = YAML.load_file('config.yaml')


# ruby 2.2.2
OpenSSL::SSL.const_set(:VERIFY_PEER, OpenSSL::SSL::VERIFY_NONE)

# open workbook
session = GoogleDrive::Session.from_config("config.json")
ws = session.spreadsheet_by_key(config["GOOGLE_SHEET"]).worksheets[0]

# # set up config
MERRIAM_API_KEY = config["MERRIAM_API_KEY"]
PEARSON_API_KEY = config["PEARSON_API_KEY"]
WORKNIK_API_KEY = config["WORKNIK_API_KEY"]
error_words = []
fetched_words = 0
skipped_words = 0


def write_to_book(workbook, row, attributes)
  workbook[row, 2] = attributes[:chinese] || "ERROR: COULD NOT FETCH" if workbook[row, 2] == "ERROR: COULD NOT FETCH" || workbook[row, 2] == ""
  workbook[row, 3] = attributes[:part_of_speech] || "ERROR: COULD NOT FETCH" if workbook[row, 3] == "ERROR: COULD NOT FETCH" || workbook[row, 3] == ""
  workbook[row, 4] = attributes[:ipa] || "ERROR: COULD NOT FETCH" if workbook[row, 4] == "ERROR: COULD NOT FETCH" || workbook[row, 4] == ""
  workbook[row, 5] = attributes[:sentence] || "ERROR: COULD NOT FETCH" if workbook[row, 5] == "ERROR: COULD NOT FETCH" || workbook[row, 5] == ""
  workbook[row, 6] = attributes[:definition] || "ERROR: COULD NOT FETCH" if workbook[row, 6] == "ERROR: COULD NOT FETCH" || workbook[row, 6] == ""
  workbook.save
end

def write_errors(workbook, row, headword, error)
  error_words << [headword, row]
  workbook[row, 3] = "ERROR: COULD NOT FETCH" if workbook[row, 3] == ""
  workbook[row, 4] = "ERROR: COULD NOT FETCH" if workbook[row, 4] == ""
  workbook[row, 5] = "ERROR: COULD NOT FETCH" if workbook[row, 5] == ""
  workbook[row, 6] = "ERROR: COULD NOT FETCH" if workbook[row, 6] == ""
  workbook.save
  puts "ERROR SKIPPING \'#{headword}\'"
  puts error.class
  puts error.message
end

def set_attribute_hash(attribute_hash, pronunciations, correct_results, search_part_of_speech)
  attribute_hash[:ipa] = pronunciations.join(", ")
  definitions = []
  correct_results.each do |result|
    definitions << result["senses"][0]["definition"][0] if (!correct_results.nil? && 
      !correct_results[0].nil? && !correct_results[0]["senses"].nil? && 
      !correct_results[0]["senses"][0].nil? && !correct_results[0]["senses"][0]["definition"].nil?)
  end
  # attribute_hash[:definition] = correct_results[0]["senses"][0]["definition"][0] if (!correct_results.nil? && !correct_results[0].nil? && !correct_results[0]["senses"].nil? && !correct_results[0]["senses"][0].nil? && !correct_results[0]["senses"][0]["definition"].nil?)
  attribute_hash[:definition] = definitions.join("; ")

  attribute_hash[:sentence] = correct_results[0]["senses"][0]["examples"][0]["text"] if (!correct_results.nil? && !correct_results[0].nil? && !correct_results[0]["senses"][0].nil? && !correct_results[0]["senses"][0]["examples"].nil?)

  attribute_hash[:sentence] ||= correct_results[0]["senses"][0]["collocation_examples"][0]["example"]["text"] if !correct_results[0]["senses"].nil? &&
  !correct_results[0]["senses"][0].nil? && !correct_results[0]["senses"][0]["collocation_examples"].nil? && !correct_results[0]["senses"][0]["collocation_examples"][0].nil? &&
  !correct_results[0]["senses"][0]["collocation_examples"][0]["example"].nil? && !correct_results[0]["senses"][0]["collocation_examples"][0]["example"]["text"].nil?

  attribute_hash[:part_of_speech] = search_part_of_speech || correct_results[0]["part_of_speech"]
end

def merge_pronunciations(correct_results, pronunciations)
  correct_results.each do |result|
    if !result["pronunciations"].nil?
      result["pronunciations"].each do |p|
        pronunciations << p["ipa"]
      end
    end
  end
end

def print_results(fetched_words, skipped_words, error_words)
  puts "************************************"
  puts "Completed."
  puts "************************************"
  puts "Fetched: #{fetched_words} Skipped: #{skipped_words} Errors: #{error_words.length} "
  puts "************************************"
  puts "ERRORS:"
  puts "************************************"
  error_words.each do |error_word|
    puts "#{error_word[0]} on line #{error_word[1]}"
  end
end

def create_html_file(words)
  fileHtml = File.new("failures.html", "w+")
  fileHtml.puts "<HTML>"
  fileHtml.puts "<HEAD>Failures</HEAD>"
  fileHtml.puts "<BODY>"
  fileHtml.puts "<h1>Failures</h1>"
  fileHtml.puts "<ol>"
  words.each do |word|
    fileHtml.puts "<li> <a href='http://learnersdictionary.com/definition/#{word[0]}' target='_blank'> row: #{word[1]} : #{word[0]} </a></li>"
  end
  fileHtml.puts "</ol>"
  fileHtml.puts "</BODY>"
  fileHtml.puts "</HTML>"
  fileHtml.close()
end

def find_ipa
  agent = Mechanize.new
  page = agent.get('http://lingorado.com/ipa/')
  query = "dog"
  form = page.forms[0]
  form["text_to_transcribe"] = query
  page = agent.submit(form)
  output = page.parser.css(".transcribed_word")[0].text
end



# loop through all rows except the header
(2..ws.num_rows).each do |row|
  # ws.num_rows
  headword =  ws[row, 1]

  # If row already completed skip this row
  if false
    skipped_words += 1
    puts "\'#{headword}\' already completed" 
  else

  search_part_of_speech = ws[row, 3]
  search_part_of_speech = false if search_part_of_speech == "" || search_part_of_speech == "ERROR: COULD NOT FETCH"
  attribute_hash = {}

    begin
      # response_1
      chinese_learner_dictionary_url = "http://api.pearson.com/v2/dictionaries/ldec/entries?headword=#{headword}"
      chinese_learner_dictionary_url << "&part_of_speech=#{search_part_of_speech}" if search_part_of_speech
      response_1 = HTTParty.get(chinese_learner_dictionary_url)
      # skip if there's a problem
      unless response_1["status"] != 200 || response_1["results"].length <= 0
        results = response_1["results"]
        correct_results = []
        if search_part_of_speech
          correct_results = results.select { |d| d["headword"] == headword && d["part_of_speech"] == search_part_of_speech }
        else
          correct_results = results.select { |d| d["headword"] == headword }
        end
        attribute_hash[:chinese] = correct_results[0]["senses"][0]["translation"] if (!correct_results[0].nil? && !correct_results[0]["senses"].nil? && !correct_results[0]["senses"][0].nil?)
      end #response 1

      # response_2
      learner_dictionary_url = "http://api.pearson.com/v2/dictionaries/lasde/entries?headword=#{headword}"
      learner_dictionary_url << "&part_of_speech=#{search_part_of_speech}" if search_part_of_speech
      response_2 = HTTParty.get(learner_dictionary_url)
      # skip if there's a problem
      unless response_2["status"] != 200 || response_2["results"].length <= 0
        results = response_2["results"]
        pronunciations = []
        if search_part_of_speech
          correct_results = results.select { |d| d["headword"] == headword && d["part_of_speech"] == search_part_of_speech }
        else
          correct_results = results.select { |d| d["headword"] == headword }
        end
        merge_pronunciations(correct_results, pronunciations)
        set_attribute_hash(attribute_hash, pronunciations, correct_results, search_part_of_speech)
      end # response_2

      # response_3
      if attribute_hash[:ipa] == "" || attribute_hash[:ipa].nil? || attribute_hash[:ipa] == "" || attribute_hash[:ipa].nil? || attribute_hash[:part_of_speech] == "" || attribute_hash[:part_of_speech].nil?

        response_failed = false
        merriam_url = "http://www.dictionaryapi.com/api/v1/references/learners/xml/#{headword}?key=#{MERRIAM_API_KEY}"

        begin
          response_3 = HTTParty.get(merriam_url)
          rescue Exception => error
            response_failed = true
        end

        if !response_failed
          response_3_hash =  Hash.from_xml(response_3)
          entry = response_3_hash["entry_list"]["entry"][0]
          ipa_2 = entry["pr"]
          def_2 = entry["def"]["dt"][0].gsub(":", "")
          pos_2 = entry["fl"]
          attribute_hash[:ipa] = ipa_2 if attribute_hash[:ipa] == "" || attribute_hash[:ipa].nil?
          attribute_hash[:definition] = def_2 if attribute_hash[:ipa] == "" || attribute_hash[:ipa].nil?
          attribute_hash[:part_of_speech] = pos_2 if attribute_hash[:part_of_speech] == "" || attribute_hash[:part_of_speech].nil?
        end
      end # response_3


        error_words << [headword, row] if attribute_hash[:ipa].nil? || attribute_hash[:definition].nil? || attribute_hash[:part_of_speech].nil? || attribute_hash[:sentence].nil?

        # write to workbook
        write_to_book(ws, row, attribute_hash)
        puts "\'#{headword}\' \u2713"
        fetched_words += 1


      # If response is bad note in worksheet
      rescue Exception => error
        error_words << [headword, row]
        ws[row, 3] = "ERROR: COULD NOT FETCH" if ws[row, 3] == ""
        ws[row, 4] = "ERROR: COULD NOT FETCH" if ws[row, 4] == ""
        ws[row, 5] = "ERROR: COULD NOT FETCH" if ws[row, 5] == ""
        ws[row, 6] = "ERROR: COULD NOT FETCH" if ws[row, 6] == ""
        ws.save
        puts "ERROR SKIPPING \'#{headword}\'"
        puts error.class
        puts error.message
    end # begin

    # If IPA isn't fetched try other site
    begin
      if attribute_hash[:ipa].nil? || attribute_hash[:ipa] == ""
        agent = Mechanize.new
        page = agent.get('http://lingorado.com/ipa/')
        query = headword
        form = page.forms[0]
        form["text_to_transcribe"] = query
        page = agent.submit(form)
        output = []
        page.parser.css(".transcribed_word").each { |span| output << span.text }
        attribute_hash[:ipa] = output.join(" ")
        ws[row, 4] = attribute_hash[:ipa]
        ws.save
      end
    rescue Exception => error
      puts "Error with fetching second IPA site"
      puts error.message
    end

  end # if
  puts "Done with row: #{row}"
  puts ((row / ws.num_rows.to_f) * 100.0).round(3).to_s + "% done #{ws.num_rows - row} left to go!"
end #each block

puts "************************************"
  puts "Completed."
  puts "************************************"
  puts "Fetched: #{fetched_words} Skipped: #{skipped_words} Errors: #{error_words.length} "
  puts "************************************"
  puts "ERRORS:"
  puts "************************************"
  create_html_file(error_words)
  error_words.each do |error_word|
    puts "#{error_word[0]} on line #{error_word[1]}"
  end



