dim version
version = "2.2"

Sub setup
  Assert.Trace "version:" & "ruby.object." & version
End Sub

Sub TestRubyizeVersion
  set r = CreateObject("ruby.object." & version)
  Assert.Equals version & ".0", r.Version
End Sub

Sub TestRubyVersion
  set r = CreateObject("ruby.object." & version)
  Assert.IsSomething r, "can't create ruby.object"
  Assert.Equals version, Left(r.RubyVersion, Len(version))
End Sub

Sub TestString
  set r = CreateObject("ruby.object." & version)
  set str = r.rubyize("abcdef")
  Assert.Equals "ABCDEF", str.upcase
End Sub

Sub TestNum
  set r = CreateObject("ruby.object." & version)
  set num = r.rubyize(12345)
  Assert.Equals 12346, num.next
End Sub

Sub TestRegex
  set r = CreateObject("ruby.object." & version)
  set reg = r.erubyize("/\Azb(\d+)(a?)C/")
  Assert.IsSomething reg
  Assert.Trace reg.to_s
  Set m = reg.match("zb321abC")
  Assert.IsNothing m
  Set m = reg.match("zb321aC")
  Assert.IsSomething m
  a = m.to_a
  Assert.Equals "zb321aC", a(0)
  Assert.Equals "321", a(1)
  Assert.Equals "a", a(2)
End Sub

Sub TestClassDef
  set r = CreateObject("ruby.object." & version)
  set obj = r.erubyize("class X;def test(x);x + 8;end;end;X.new")
  x = obj.test(2)
  Assert.Equals 10, x
End Sub

Sub TestSyntaxError
  Assert.ExpectError "syntax error"
  set r = CreateObject("ruby.object." & version)
  set obj = r.erubyize("class X;def test..(x)x + 8;end;end;X.new")
End Sub

Sub TestRuntimeError
  Assert.ExpectError "ZeroDivisionError"
  set r = CreateObject("ruby.object." & version)
  set obj = r.erubyize("class X;def test(x) 8 / x;end;end;X.new")
  obj.test(0)
End Sub
