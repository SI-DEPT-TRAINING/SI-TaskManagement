# -*- encoding: utf-8 -*-
class SerchOutputModel
  def initialize(name,acount, eventList)
    #氏名
    @name = name;
    #アカウント
    @acount = acount;
    #カレンダー情報リスト
    @eventList = eventList;
  end
  attr_accessor :name
  attr_accessor :acount
  attr_accessor :eventList
end