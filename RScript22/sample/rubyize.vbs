set x = CreateObject("ruby.object.2.2")
if x is nothing then
  WSH.echo "nothing"
else
  WSH.echo x.rubyversion
  WSH.echo x.version
  Set a = x.rubyize("abcdef")
  WSH.echo a.upcase
  Set a = x.erubyize("class X;def hello;'abc';end;end;X.new")
  WSH.echo a.hello
end if