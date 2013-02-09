# -*- encoding: utf-8 -*-
require "gcalapi";
require 'time';

class ScheduleModel
  # コンストラクタ。
  def initialize(googleCalObj)
    # GoogleCalendarを受け取り、クラス変数に保持する。
    @googleCalObj = googleCalObj;
  end
  attr_accessor :googleCalObj
  attr_accessor :username
  
  
  # メソッド定義。
  def getTitle()
    return googleCalObj.title
  end
  def getDesc()
    return googleCalObj.desc
  end
  def getWhere()
    return googleCalObj.where
  end
  def getWorkTimeMinuts()
    startTime = googleCalObj.st
    endTime = googleCalObj.en
    
    days = (endTime - startTime).divmod(24*60*60) #=> [2.0, 45000.0]
    hours = days[1].divmod(60*60) #=> [12.0, 1800.0]
    mins = hours[1].divmod(60) #=> [30.0, 0.0]  
    
    answer = days[0].to_i * 24*60 + hours[0].to_i * 60 + mins[0].to_i
    return answer
  end
  def getStartDate()
    return googleCalObj.st
  end
  def getEndDate()
    return googleCalObj.en
  end
  
  def isProcessTarget()
    codeRegex = "[A-Z]{0,1}[0-9]{1,3}"
    delimRegex = "[ \t]*,[ \t]*"
    allRegex = codeRegex + delimRegex + codeRegex + delimRegex + codeRegex
    desc = googleCalObj.desc
    if desc.nil? == true
      return false
    end
    if desc.scan(/#{allRegex}/).size == 0
      return false
    end
    return true
  end
end
