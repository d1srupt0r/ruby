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
      CSV.foreach arg do |row|
        # row row row your boat
        r = row.inspect if row.inspect != "nil"
        header = r =~ /(2017)/
        is_date = !(Date.parse r rescue nil).nil?
        dates = [] if header
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
              e.dtstart = Icalendar::Values::Date.new(d) if !d.nil?
              e.dtend = Icalendar::Values::Date.new(d) if !d.nil?
              e.summary = value
              #e.description = "w00t w00t celebrate!?!"
              e.ip_class = "PRIVATE"
            end
          end
        end
      end
      File.write "#{options[:user]}.ics", calendar.to_ical
    when ".xls"
    when ".xlsx"
      xlsx = Roo::Excelx.new arg
      pa "Converting #{File.basename arg}...", "afa"
      xlsx.to_csv "#{File.expand_path "~/Documents"}/#{File.basename arg, ".*"}.csv"
    end
end

pa "Finished running at: #{Time.new}", "aaa"