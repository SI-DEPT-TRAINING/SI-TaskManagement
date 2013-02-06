# -*- encoding: utf-8 -*-
require "gcalapi";

class ScheduleList 
  # コンストラクタ。
  def initialize()
    # GoogleCalendarを受け取り、クラス変数に保持する。
    @list = [];
  end
  
  
  # メソッド定義。
  def add(obj)
    @list.push(obj)
  end
  def getMemberList()
    list = []
    @list.each do |tmp|
      list.push(tmp.username)
    end
    # 
    list.uniq!
    return list
  end
  def setMemberList(list)
    @list = list
  end
  def getList()
    @list;
  end
  def narrowMember(username)
    list = []
    @list.each do |tmp|
      if tmp.username == username
        list.push(tmp)
      end
    end
    # 
    return list
  end
end
