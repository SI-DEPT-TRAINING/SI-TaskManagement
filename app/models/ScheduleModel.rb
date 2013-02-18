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
  def getDesc_FirstLine()
    str = googleCalObj.desc
    if str==nil
      return nil
    end
    # 先頭行のみを返す。
    return str[/[^\r\n]*/]
  end
  def checkSectionList?()
    sectionList = getSectionList()
    sectionList.each do |section|
      if (section == nil || section == "")
        return false
      end
    end
    return true
  end
  def getSectionList()
    line = getDesc_FirstLine()
    if line == nil
      return ["", "", ""]
    end
    /([^,]*),([^,]*),([^,]*)/ =~ getDesc_FirstLine()
    sectionList = [ $1, $2, $3 ]
    return sectionList
  end
  def getWhere()
    return googleCalObj.where
  end
  def getWorkTimeHours(unit=1)
    min = getWorkTimeMinuts(unit)
    hour = min / 60.0
    return hour
  end
  def getWorkTimeMinuts(unit=1)
    startTime = googleCalObj.st
    endTime = googleCalObj.en
    
    days = (endTime - startTime).divmod(24*60*60) #=> [2.0, 45000.0]
    hours = days[1].divmod(60*60) #=> [12.0, 1800.0]
    mins = hours[1].divmod(60) #=> [30.0, 0.0]  
    
    min = days[0].to_i * 24*60 + hours[0].to_i * 60 + mins[0].to_i
    
    mod = min % unit
    return min - mod 
    if(mod == 0)
      return min / unit
    end
    return min / unit + 1
  end
  def getStartDate()
    return googleCalObj.st
  end
  def getEndDate()
    return googleCalObj.en
  end
  
  def randBool?()
      return false
    if ( 50 < rand(100) )
      return false
    end
    return true
  end
  def isProcessTarget?()
    sectionList = getSectionList()
    if (sectionList[0] == "" || sectionList[1] == "" || sectionList[2] == "")
      return randBool?()
    end
    if (sectionList[0] == nil || sectionList[1] == nil || sectionList[2] == nil)
      return randBool?()
    end
    
    # 3セクションとも設定している。
    return true
  end
  
  def getTermString()
    # 開始と終了の日時を取得。
    startDate = @googleCalObj.st
    endDate = @googleCalObj.en
    # 開始日付, 開始時刻, 終了時刻を文字列として取得。
    start_date_string = startDate.strftime("%Y/%m/%d")
    start_time_string = startDate.strftime("%H:%M")
    end_time_string = endDate.strftime("%H:%M")
    # 期間の文字列を作成。
    ret_value = sprintf("%s  %s - %s", start_date_string, start_time_string, end_time_string)
    return ret_value
  end
  
  def getSameSectionIndexBeforeMyself(scheduleList)
    index = 0
    mySectionList = getSectionList()
    if (checkSectionList? == false)
      return -1
    end
    scheduleList.each do |schedule|
      if (schedule.isProcessTarget? == false)
        next
      end
      tmpSectionList = schedule.getSectionList()
      if (mySectionList[0] == "A001")
        i = 1
      end
      if (mySectionList[0] != tmpSectionList[0])
      index +=1
        next
      end
      if (mySectionList[1] != tmpSectionList[1])
      index +=1
        next
      end
      if (mySectionList[2] != tmpSectionList[2])
      index +=1
        next
      end
      
      # 分類の内容が一致した。
      if (self == schedule)
        return -1
      end
      if (self != schedule)
        # ただしオブジェクトは一致せず。
        return index
      end
      index +=1
    end
    
    return -1
  end
end
