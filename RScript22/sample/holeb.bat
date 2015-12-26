@echo off
"%~dp0ruby" -x "%~f0" %*
goto endofruby
#!/usr/bin/ruby
system("start mshta \"file:///#{File.dirname($0)}/holeb.hta\"")
__END__
:endofruby
