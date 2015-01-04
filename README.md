RScript22
=========

IActiveScript implementation for Ruby2.x both Win32/Win64 MT

* IActiveScript interface for ruby interpreter
  * It can write ruby code as embeded script. For example HTA (Html application).
* Rubyizer
  * Rubyizer is simple COM object that exports ruby object into the caller's environment.

# Rubyizer
## Rubyize example

The below VBScript code illustrates the power of Rubyizer.

```
'test.vbs
set r = CreateObject("ruby.object.2.2")
WSH.echo r.RubyVersion  '=>2.2.0-p0 (x64-mswin64_100)
Set reg = r.erubyize("/\Azb(\d+)(a?)C/")
Set m = reg.match("zb321aC")
If Not m Is Nothing Then
  group = m.to_a
  WSH.echo "match:" & group(0)      '=>zb321aC
  WSH.echo "1st group:" & group(1)  '=>321
  WSH.echo "2nd group:" & group(2)  '=>a
End If
```

### description

line 2 : instantiate Rubyizer
line 4 : Rubyizer#erubyize is evaluate the string argument as Ruby script and returns the evaluated value as an object.
         In this sample code, the argument is a ruby's Regexp literal, therefore the returned value is Regexp object.
line 5 : send 'match' to the Regex object and set the result object to variable 'm'.
line 7 : get the result array from Match object and set it to variable 'group'.
line 8 : Because 'group' is VBScript's array, it can accessed its element with (index) syntax.

## Rubyizer methods

**Rubyizer#rubyize**

convert the argument to the rubyized object.

```
set r = CreateObject("ruby.object.2.2")
set num = r.rubyize(3)  ' get ruby Fixnum object
WSH.echo num.next       ' => 4 
```

**Rubyizer#erubyize**

evaluated the argument as ruby code. it returns the evaluated result as a rubyized object.

