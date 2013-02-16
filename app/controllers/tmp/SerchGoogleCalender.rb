# -*- encoding: utf-8 -*-
require "kconv";
require "gcalapi";
require './model/SerchOutputModel.rb'

=begin 
Googleカレンダー情報取得クラス
=end
class SerchGoogleCalender
  
  def initialize(serchCondtionModel)
    @serchCondtionModel = serchCondtionModel;
  end

=begin 
Googleカレンダー情報取得
=end
  def getGoogleCalenderData()
    #Serviceオブジェクトの作成
    srv = GoogleCalendar::Service.new(@serchCondtionModel.masterAcount, @serchCondtionModel.pass)
    serchOutputModelList = Array.new{};
    @serchCondtionModel.acountListMap.each do |name,acount|
      # 非公開URL（データの追加/更新/削除が可能）
      feed = "http://www.google.com/calendar/feeds/" + acount + "/private/full"
      #Calendarオブジェクトの作成
      cal = GoogleCalendar::Calendar.new(srv, feed)
      #検索条件を設定しデータ取得実行
      eventList = cal.events(:'max-results' => @serchCondtionModel.maxResults ,
                             :'start-min'   => @serchCondtionModel.startMin   ,
                             :'start-max'   => @serchCondtionModel.startMax   ,
                             :orderby       => @serchCondtionModel.orderBy    ,
                             :sortorder     => @serchCondtionModel.sortOrder  )
      serchOutputModel = SerchOutputModel.new(name,acount,eventList)
      serchOutputModelList <<  serchOutputModel;
    end
    return serchOutputModelList
  end
end
