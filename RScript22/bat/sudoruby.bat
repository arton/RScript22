@echo off
"%~dp0ruby" -x "%~f0" %*
@goto endofruby
#!/bin/ruby
require 'suexec'
path = File::dirname($0).gsub(File::SEPARATOR, File::ALT_SEPARATOR)
ENV['PATH'] = "#{path};#{ENV['PATH']}"
if ARGV.size > 0
  ARGV[0] = File.expand_path(ARGV[0])
end
SuExec.exec 'cmd', *['/C', %|"#{File::dirname($0)}\\setenvpath.bat"|, *ARGV]
__END__
:endofruby
