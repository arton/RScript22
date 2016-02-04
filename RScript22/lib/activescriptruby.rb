# coding: utf-8
#
# Copyright(c) 2004,2015,2016 arton
#
# This library is free software; you can redistribute it and/or
# modify it under the terms of the GNU Lesser General Public
# License as published by the Free Software Foundation; either
# version 2.1 of the License, or (at your option) any later version.
#
# This library is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
# Lesser General Public License for more details.
#

require 'win32ole'
require 'thread'
require 'set'

module ActiveScriptRuby

  class AsrProxy
    USER_DISPID = 28000
    alias :include :extend
    def initialize(obj)
      @target = obj
      @dispid = { }
      @method_name = { }
      @current_dispid = USER_DISPID
    end
    alias :org_respond_to? :respond_to?
    def respond_to?(m)
      if org_respond_to?(m)
        true
      else
        @target.respond_to?(m)
      end
    end
    def to_enum
      @target.to_enum
    end
    def add_method(name, script, fn, line)
      define_singleton_method(name.to_sym) {
        instance_eval(script, fn, line)
      }
    end
    def call(dispid, *args)
      @target.__send__(@method_name[dispid], *args) if @target
    end
    def get_method_id(name)
      normalized = name.upcase.intern
      val = @dispid[normalized]
      unless val
        method_name = @target.methods.find { |x| x.casecmp(normalized) == 0 }
        if method_name
          @dispid[normalized] = val = @current_dispid += 1
          @method_name[val] = method_name
        else
          val = -1
        end
      end
      val
    end
  end

  class SelfProxy < AsrProxy
    def initialize
      super(self)
    end
  end

  class GlobalProxy < AsrProxy
    @lock = Mutex.new

    NamedItem = Struct.new(:obj, :global) do
      def eql? (other) obj.eql?(other.obj); end
      def hash; obj.hash; end
    end

    def initialize(bridge)
      super(self)
      @bridge = bridge
      @named_items = {}
      @global = nil
    end

    def to_variant(obj)
      @bridge.PassObject(obj)
    end

    def self_to_variant
      @bridge.PassObject(self)
    end

    def to_proxy(val)
      obj = AsrProxy.new(val)
      @bridge.PassObject(obj)
      obj
    end

    def ruby_version
      "#{RUBY_VERSION}-p#{RUBY_PATCHLEVEL} (#{RUBY_PLATFORM})"
    end

    def create_event_handler(itemname, procname, code, line)
      val = @named_items[itemname]
      unless val
        obj = SelfProxy.new
        add_named_item(itemname, obj, false)
      else
        if val.global
          obj = self
        else
          obj = val.obj
        end
      end
      obj.add_method(procname, code, '(asr)', line)
      obj
    end

    def create_event_proc(itemname, code, line)
      val = @named_items[itemname]
      unless val
        obj = SelfProxy.new
        add_named_item(itemname, obj, false)
      else
        if val.global
          obj = self
        else
          obj = val.obj
        end
      end
      Proc.new do
        obj.instance_eval(code, '(asr)', line)
      end
    end

    def method_missing(id, *args)
      @global.__send__(id, *args) if @global
    end

    def add_named_item(name, obj, global)
      @named_items[name] = NamedItem.new(obj, global)
      if name[0].between?('A', 'Z')
        GlobalProxy.const_set(name, obj)
      else
        define_singleton_method(name.to_sym) { obj }
      end
      if global
        @global = obj
      end
      # ASR 1.0 compatibility
      instance_eval("@#{name} = #{name}")
      GlobalProxy.const_set(name.capitalize, obj)
    end

    def add_constant(name, variant)
      define_singleton_method(name.to_sym) { variant.value }
    end

    def remove_item(name)
      val = @named_items[name]
      if val
        if name[0].between?('A', 'Z')
          GlobalProxy.delete_constant(name.to_sym)
        else
          define_singleton_method(name.to_sym) { nil }
        end
        val.obj.ole_free if WIN32OLE === val.obj
        if val.global
          @global = nil
        end
        # ASR 1.0 compatibility
        instance_eval("@#{name} = nil")
        GlobalProxy.delete_constant(name.capitalize)
        @named_items[name] = nil
      end
    end

    def self.delete_constant(name)
      remove_const(name)
    end

    def remove_items
      @named_items.each { |key, val| remove_item(key) if val }
      @named_items.clear
    end
  end
end
