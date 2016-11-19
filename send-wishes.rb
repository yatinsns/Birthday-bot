#!/usr/bin/env ruby

require 'roo'

RIVA_WEBHOOK = "<WEBHOOK>"
ZETA_WEBHOOK = "<WEBHOOK>"

def select_birthdays_today_in_xlsx(filename, sheet_index)
  xlsx = Roo::Excelx.new(filename)

  current_date = Date.parse(Time.now.to_s)

  xlsx.sheet(sheet_index).select do |row|
    if row[3].is_a?(Date)
      birthday = row[3]
      birthday.month == current_date.month && birthday.day == current_date.day  
    end
  end.map do |row|
    row[2]
  end
end

def get_random_birthday_wish_text
  birthday_wishes = File.read('birthday_wishes_list.txt').split("\n")
  birthday_wishes[rand birthday_wishes.length]
end

def wish person_name
  "Hey #{person_name},\n#{get_random_birthday_wish_text}"
end

def send_text(webhook, text)
  `curl -X POST -d '{"text":"#{text}"}' -H "Content-Type:application/json;charset=UTF-8" #{webhook}`
end

def send_wishes
  riva_names = select_birthdays_today_in_xlsx "./birthdays.xlsx", 0
  riva_names.map do |name|
    wish name
  end.each do |text|
    send_text RIVA_WEBHOOK, text
  end

  zeta_names = select_birthdays_today_in_xlsx "./birthdays.xlsx", 1
  zeta_names.map do |name|
    wish name
  end.each do |text|
    send_text ZETA_WEBHOOK, text
  end
end

send_wishes
