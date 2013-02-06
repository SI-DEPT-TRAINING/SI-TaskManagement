# -*- encoding: utf-8 -*-
require "gcalapi";

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
  def getEndDate2()
    return googleCalObj.en
  end
  def getStartDate()
    return googleCalObj.st
  end
  def getEndDate()
    return googleCalObj.en
  end
  
  def isProcessTarget()
    codeRegex = "[A-Z]{1}[0-9]{3}"
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
