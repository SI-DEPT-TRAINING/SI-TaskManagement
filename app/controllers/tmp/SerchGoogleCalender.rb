# -*- encoding: utf-8 -*-
require "kconv";
require "gcalapi";
require './app/controllers/tmp/SerchOutputModel.rb'

=begin 
Google繧ｫ繝ｬ繝ｳ繝��諠��蜿門ｾ励け繝ｩ繧ｹ
=end
class SerchGoogleCalender
  
  def initialize(serchCondtionModel)
    @serchCondtionModel = serchCondtionModel;
  end

=begin 
Google繧ｫ繝ｬ繝ｳ繝��諠��蜿門ｾ�
=end
  def getGoogleCalenderData()
    #Service繧ｪ繝悶ず繧ｧ繧ｯ繝医�菴懈�
    srv = GoogleCalendar::Service.new(@serchCondtionModel.masterAcount, @serchCondtionModel.pass)
    serchOutputModelList = Array.new{};
    @serchCondtionModel.acountListMap.each do |name,acount|
      # 髱槫�髢偽RL�医ョ繝ｼ繧ｿ縺ｮ霑ｽ蜉�譖ｴ譁ｰ/蜑企勁縺悟庄閭ｽ��
      feed = "http://www.google.com/calendar/feeds/" + acount + "/private/full"
      #Calendar繧ｪ繝悶ず繧ｧ繧ｯ繝医�菴懈�
      cal = GoogleCalendar::Calendar.new(srv, feed)
      #讀懃ｴ｢譚｡莉ｶ繧定ｨｭ螳壹＠繝��繧ｿ蜿門ｾ怜ｮ溯｡�
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
