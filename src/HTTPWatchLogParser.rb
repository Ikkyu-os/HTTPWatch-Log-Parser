# Imports
require 'win32ole'
require 'csv'


# Initialize all objects
controller = WIN32OLE.new('HttpWatch.Controller')

# get all files to parse
files = Dir.glob("*.hwl")

CSV.open("D:\\Documents\\Projects\\CUI\\ARD\\SideProjects\\HTTPWatchParser\\RubyScriptParsing\\results\\outputLogs.csv", "wb") do |csv|
  csv << ["Network", "Blocked", "DNSLookup", "Connect", "Wait","Receive","URL", "File","JSession"]
  files.each do |file|
    puts "Processing #{file} ..."
    log = controller.OpenLog(file)
    log.Entries.each do |entry|
      if /app_ard/.match(entry.URL) && entry.Timings.CacheRead.nil? == false
        urlData = entry.URL.split(';');
        csv << [entry.Timings.Network.Duration, entry.Timings.Blocked.Duration, entry.Timings.DNSLookup.Duration, entry.Timings.Connect.Duration, entry.Timings.Wait.Duration, entry.Timings.Receive.Duration, urlData[0], file, urlData[1]]
      end
    end
  end
end
