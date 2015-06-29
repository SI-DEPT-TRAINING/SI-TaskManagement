# -*- encoding: utf-8 -*-
#require "gcalapi";

class ScheduleList 
  # コンストラクタ。
  def initialize()
    # GoogleCalendarを受け取り、クラス変数に保持する。
    @list = [];
  end

  # メソッド定義。 kogane
  # メソッド定義。 kogane kogane kogane kogane kogane
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
    # list = []
    list = ScheduleList.new()
    @list.each do |tmp|
      if tmp.username == username
        # list.push(tmp)
        list.add(tmp)
      end
    end
    # 
    return list
  end
  
  def chk(list, schedule)
    list.each do |tmp|
      if (schedule.isSameSection?(tmp) == true)
        # 一致するセクションがlist内に存在する。
        return true
      end
    end
    # 一致するセクションがlist内に存在しない。
    return false
  end
  def getMargeSectionIndex(argSchedule)
    # まずはセクションが正しく入力しているものをリスト内に絞り込む。
    tmpList = Array.new()
    @list.each do |schedule|
      # 処理対象となる正常なスケジュール（セクションが正しく入力）のみ取得
      if (schedule.isProcessTarget? == true)
        # 自分と同一のセクションのスケジュールがあるなら追加しない。なければ追加する。
        if (chk(tmpList, schedule) == false)
          tmpList.push(schedule)
        end
      end
    end
    
    # tmpList内のインデックスを返す。
    idx = 0
    tmpList.each do |tmp|
      if (tmp.isSameSection?(argSchedule))
        # セクションが一致するなら現在のインデックスを返す。
        return idx
      end
      idx = idx + 1
    end
    
    # 一致するセクションがなかった。
    return -1
  end
  def getSameSectionBeforeMyself?(argSchedule)
    # 引数のスケジュールが保持するセクションと同一のものならそのスケジュールを返す。
    argSectionList = argSchedule.getSectionList()
    if (argSchedule.checkSectionList? == false)
      return nil
    end
    
    @list.each do |schedule|
      if (schedule.isProcessTarget? == false)
        next
      end
      
      tmpSectionList = schedule.getSectionList()
      if (argSectionList[0] != tmpSectionList[0])
        next
      end
      if (argSectionList[1] != tmpSectionList[1])
        next
      end
      if (argSectionList[2] != tmpSectionList[2])
        next
      end
      
      # 分類の内容が一致した。
      return schedule
      # index +=1
    end
    
    # 一致する分類の内容なし。
    return nil
  end
end
