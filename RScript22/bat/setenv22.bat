@echo off
@if not "%~d0" == "~d0" goto WinNT
\usr\bin\ruby -x "/usr/bin/setenv22.bat" %1 %2 %3 %4 %5 %6 %7 %8 %9
@goto endofruby
:WinNT
"%~dp0ruby" -x "%~f0" %*
@goto endofruby
#!/usr/bin/ruby
require 'winpath'
path = Pathname.new(File.dirname($0)).shortname.gsub(File::SEPARATOR, File::ALT_SEPARATOR)
ENV['PATH'] = "#{path};#{ENV['PATH']}"
if ARGV.size > 0
  cmd = 'start ruby "' + ARGV[0] + '"'
else
  if ENV['OS'] == 'Windows_NT'
    cmd = "start cmd"
  else
    cmd = "start command"
  end
end
system(cmd)
__END__
:endofruby
