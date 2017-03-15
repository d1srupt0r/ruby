#!/usr/bin/ruby
require 'csv'
require 'etc'
require 'icalendar'
require 'optparse'
require 'paint/pa'
require 'roo'

calendar = Icalendar::Calendar.new
days = ["","THURSDAY","FRIDAY","SATURDAY","SUNDAY","MONDAY","TUESDAY","WEDNESDAY"]
now = Time.new.strftime '"%-d-%b'
options = { :user => Etc.getlogin }

OptionParser.new do |opts|
  opts.banner = "Usage: schedule.rb [options]"
  opts.on("-uUSER", "--user=USER", "User") do |user|
    options[:user] = user
  end
  opts.on("-v", "--[no-]verbose", "Show mour info") do
    options[:verbose] = true
  end
  opts.on("-h", "--help", "Display help") do
    pa opts, "aaa"
    exit
  end
end.parse!

ARGV.each do |arg|
  next if !File.file? arg
  case File.extname(arg)
    when ".csv"
      dates = []
      # loop through each file
      CSV.foreach arg do |row|
        # row row row your boat
        r = row.inspect if row.inspect != "nil"
        header = r =~ /(2017)/
        dates = [] if header
        is_date = !(Date.parse r rescue nil).nil?
        # ToDo: don't filter out values, simply handle them differently
        next unless header or r.downcase.include? options[:user]
        pa r, "bbb" if header and options[:verbose]
        # loop once per weekday Mon-Fri
        (1..7).each do |i|
          value = row[i].inspect.gsub /"/,""
          if is_date
            dates.push Date.parse "#{value}-17"
          else
            color = (["nil", "off", "available", "trip??", "?"].include? value.downcase) ? "555" : "afa";
            pa "#{i}) #{dates[i - 1]} #{days[i]}: #{value}", color if options[:verbose]
            next if color == "555"
            calendar.event do |e|
              d = dates[i - 1]
              e.description = ""
              e.dtend = Icalendar::Values::Date.new(d) if !d.nil?
              e.dtstart = Icalendar::Values::Date.new(d) if !d.nil?
              e.ip_class = "PRIVATE"
              e.location = ""
              e.summary = value
            end
          end
        end
      end
      File.write "#{File.expand_path "~/Documents"}/#{options[:user]}.ics", calendar.to_ical
    when ".xls"
    when ".xlsx"
      xlsx = Roo::Excelx.new arg
      pa "Converting #{File.basename arg}...", "afa"
      xlsx.to_csv "#{File.expand_path "~/Documents"}/#{File.basename arg, ".*"}.csv"
    when ".ics"
      events = Icalendar::Event.parse File.open arg
      events.each do |event|
        pa "#{event.dtstart} #{event.summary}", "afa" if options[:verbose]
      end
      pa "Found #{events.length} events", "aaa"
    end
end

pa "Finished running at: #{Time.new}", "aaa"