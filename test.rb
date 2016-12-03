require 'mechanize'
require "pry"
agent = Mechanize.new
page = agent.get('http://lingorado.com/ipa/')

query = "dog"
form = page.forms[0]
form["text_to_transcribe"] = query
page = agent.submit(form)
output = page.parser.css(".transcribed_word")[0].text

puts "============"
puts output
puts "============"
