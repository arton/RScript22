@echo off
@if not "%~d0" == "~d0" goto WinNT
\bin\ruby -x "/bin/make_bat.bat" %1 %2 %3 %4 %5 %6 %7 %8 %9
@goto endofruby
:WinNT
"%~dp0ruby" -x "%~f0" %*
@goto endofrubybat
#!/bin/ruby
if ARGV.length == 0 || ARGV.length > 2
  $stderr.puts 'usage: make_bat script [dest-dir = %ruby-bindir%]'
  exit(1)
end
if ARGV.length == 2
  path = ARGV[1].gsub(File::ALT_SEPARATOR, File::SEPARATOR)
else
  path = File.dirname($0)
end
dest = "#{path}/#{File.basename(ARGV[0], '.*')}.bat"
File.open(dest, 'w') do |out|
  out.write <<HEAD
@echo off
@if not "%~d0" == "~d0" goto WinNT
#{File.dirname($0).gsub(File::SEPARATOR, File::ALT_SEPARATOR)}#{File::ALT_SEPARATOR}ruby -x "#{dest}" %1 %2 %3 %4 %5 %6 %7 %8 %9
@goto endofruby
:WinNT
"%~dp0ruby" -x "%~f0" %*
@goto endofruby

HEAD
  File.open(ARGV[0], 'r').each_line do |line|
    out.puts line
  end.close
  out.puts '__END__'
  out.puts ':endofruby'
end
__END__
:endofrubybat
