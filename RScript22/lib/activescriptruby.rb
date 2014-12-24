# coding: utf-8
#
#  Copyright(c) 2014 arton
#

require 'win32ole'
require 'thread'

class ActiveScriptRuby
  @lock = Mutex.new
  @instances = {}
  NamedItem = Struct.new(:obj, :global) do
    def eql? (other) obj.eql?(other.obj); end
    def hash; obj.hash; end
  end
  
  def initialize(bridge)
    @bridge = bridge
    @named_items = {}
  end
  
  def to_variant(obj)
    @bridge.PassObject(obj)
  end
  
  def self_to_variant
    @bridge.PassObject(self)
  end

  alias :rubyize :to_variant
  
  def ruby_version
    RUBY_VERSION
  end
  
  def erubyize(str)
    to_variant(instance_eval(str))
  end
  
  def add_named_item(name, obj, global)
    @named_items[name] = NamedItem.new(obj, global)
    define_singleton_method(name.to_sym) { obj }
    if global
      # ASR 1.0 compatibility
      instance_eval("#{name.capitalize} = #{name}; @#{name} = #{name}")
    end
  end
  
  def add_constant(name, variant)
    define_singleton_method(name.to_sym) { variant.value }
  end
  
  def add_method(name, script, fn, line)
    define_singleton_method(name.to_sym) { 
      instance_eval(script, fn, line)
    }
  end
  
  def remove_item(name)
    val = @named_items[name]
    if val
      define_singleton_method(name.to_sym) { nil }
      val.obj.ole_free
      if val.global
        # ASR 1.0 compatibility
        instance_eval("#{name.capitalize} = @#{name} = nil")
      end
      @named_items[name] = nil
    end
  end
  
  def remove_items
    @named_items.each { |key, val| remove_item(key) if val }
    @named_items.clear
  end
end
