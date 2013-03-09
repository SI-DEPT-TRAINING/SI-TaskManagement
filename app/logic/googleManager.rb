# -*- encoding: utf-8 -*-
############################
# api_client OAuth2.0      #
############################
module GoogleManager
require "rubygems"
require "google/api_client"
require "yaml"
require "time"


  ############################
  # Calendar Version 3       #
  ############################
  class Calender

    #GoogleToken
    @apiClient = nil
    @calnder = nil

    #コンストラクタ
    def initialize(client)
       @apiClient = client
       @calnder = client.discovered_api('calendar', 'v3')
    end

    #カレンダー情報返却
    def getEventList(searchModel)
      
      tempMin = searchModel.startMin.gsub("/", "")
      tempMax = searchModel.startMax.gsub("/", "")
      orderBy = searchModel.orderBy
      sortOrder = searchModel.sortOrder
      maxResults = searchModel.maxResults
      min = Time.utc(tempMin[0, 4].to_i, tempMin[4,2].to_i, tempMin[6,2].to_i).iso8601
      max = Time.utc(tempMax[0, 4].to_i, tempMax[4,2].to_i, tempMax[6,2].to_i).iso8601

      calResultList = Array.new{};
      searchModel.acountListModel.each do | model |
        result =  @apiClient.execute(:api_method => @calnder.events.list,
                                     :parameters => {"calendarId" => model.acount,
                                                     "timeMin"    => min,
                                                      "timeMax"   => max,
                                                      "orderby"   => orderBy,
                                                      "sortOrder" => sortOrder,
                                                      "maxResults" => maxResults,
                                                      })
        events = result.data.items
        calResult = CalResult.new(model.name, model.acount, events)
        calResultList <<  calResult;
      end
      return calResultList;
    end

    #カレンダー情報操作メソッド
    def insertEventList
    end

    def upEventList
    end

    def deleteEventList
    end
    
  end

  #GoogleCalenderApi検索条件モデル
  class CalSearchModel
    def initialize
      # パスワード
      @pass = nil;
      # カレンダー情報取得アカウント
      @masterAcount = nil;
      # アカウントリスト
      @acountListModel = nil;
      #最大取得件数
      @maxResults = nil;
      #開始日
      @startMin = nil;
      #終了日
      @startMax = nil;
      #ソート条件
      @orderBy = nil;
      #ソート順
      @sortOrder = nil;
    end
    attr_accessor :pass
    attr_accessor :masterAcount
    attr_accessor :acountListModel
    attr_accessor :maxResults
    attr_accessor :startMin
    attr_accessor :startMax
    attr_accessor :orderBy
    attr_accessor :sortOrder
  end

  #GoogleCalenderApi検索結果モデル
  class CalResult
    def initialize(name, acount, eventList)
      #氏名
      @name = name;
      #アカウント
      @acont = acount;
      #カレンダー情報リスト
      @eventList = eventList;
    end
    attr_accessor :name
    attr_accessor :acont
    attr_accessor :eventList
  end

end