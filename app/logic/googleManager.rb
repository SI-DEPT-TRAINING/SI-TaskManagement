
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

    #定義ファイルより取得する各種パラメータ郡
    SYSTEM_YAML = YAML.load_file(File.dirname(__FILE__) + '/../../config/googleSystemProp.yml')
    ORDER_BY = SYSTEM_YAML["cal_api_orderBy"]
    SORT_ORDER = SYSTEM_YAML["cal_api_sortOrder"]
    MAX_RESULT = SYSTEM_YAML["cal_api_maxResults"]
    SINGLE_EVENTS = SYSTEM_YAML["cal_api_single_events"]
    TIMEZONE = SYSTEM_YAML["cal_api_timezone"]
    FIELDS = SYSTEM_YAML["cal_api_fields"]

    #コンストラクタ
    def initialize(client)
       @apiClient = client
       @calnder = client.discovered_api('calendar', 'v3')
    end

    #カレンダー情報返却
    def getEventList(searchModel)

      timeMin = creatTimeForUTC(searchModel.startMin)
      timeMax = creatTimeForUTC(searchModel.startMax)

      calResultList = Array.new{};
      searchModel.acountListModel.each do | model |

        begin
          result =  @apiClient.execute(:api_method => @calnder.events.list,
                                       :parameters => {"calendarId"   => model.acount,
                                                       "timeMin"      => timeMin,
                                                       "timeMax"      => timeMax,
                                                       "orderby"      => ORDER_BY,
                                                       "sortOrder"    => SORT_ORDER,
                                                       "maxResults"   => MAX_RESULT,
                                                       "singleEvents" => SINGLE_EVENTS,
                                                       "timeZone"     => TIMEZONE,
                                                       "fields"       => FIELDS
                                                        })
        rescue => exception
         logger.error "CalenderAPI ERROR: #{exception.message}"
         parameters = {:calendarId   => model.acount,
                       :timeMin      => timeMin,
                       :timeMax      => timeMax,
                       :orderby      => ORDER_BY,
                       :sortOrder    => SORT_ORDER,
                       :maxResults   => MAX_RESULT,
                       :singleEvents => SINGLE_EVENTS,
                       :timeZone     => TIMEZONE,
                       :fields       => FIELDS}
         logger.error "parameters: #{parameters.inspect}"
        end
        calResultList << creatEventsResult(result, model);
      end
      return calResultList
    end

    #カレンダーレスポンスモデルを返します
    private
    def creatEventsResult(result, model)
        events = nil
        error = nil
        if result.data != nil then
          events = result.data.items
          if result.data['error'] != nil then
            error = result.data['error']
          end
        end
        return CalResult.new(model.name, model.acount, events, error)
    end

    #UTC時間を返します
    private
    def creatTimeForUTC(time)
        tempTime = time.gsub("/", "")
        return Time.utc(tempTime[0, 4].to_i, tempTime[4,2].to_i, tempTime[6,2].to_i).iso8601
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
      # アカウントリスト
      @acountListModel = nil;
      #開始日
      @startMin = nil;
      #終了日
      @startMax = nil;
    end
    attr_accessor :acountListModel
    attr_accessor :startMin
    attr_accessor :startMax
  end

  #GoogleCalenderApi検索結果モデル
  class CalResult
    def initialize(name, acount, eventLis, errort)
      #氏名
      @name = name;
      #アカウント
      @acont = acount;
      #カレンダー情報リスト
      @eventList = eventList;
      #カレンダーエラー情報
      @error = error
    end
    attr_accessor :name
    attr_accessor :acont
    attr_accessor :eventList
    attr_accessor :error
  end

end