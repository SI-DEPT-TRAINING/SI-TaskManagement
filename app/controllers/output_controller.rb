# -*- coding: utf-8 -*-
require './app/models/CreateWorkbook.rb'
require './app/controllers/tmp/SerchCondtionModel.rb'
require './app/controllers/tmp/SerchGoogleCalender.rb'
require './app/models/ScheduleList.rb'
require './app/models/ScheduleModel.rb'

class OutputController < ApplicationController

  def output
    condi = getCondi();
    xxxGoogleEventList = getXXXXGoogleAPIEventList(condi, 'fileName')
    
    
    #カレンダー情報取得クラスオブジェクト生成
    #serchGoogleCalender = SerchGoogleCalender.new(serchCondtionModel);
    #カレンダー情報取得
    #serchOutputModelList = serchGoogleCalender.getGoogleCalenderData();
    #serchOutputModelList= Array.new();
    
    #serchOutputModelList= getXXXX(serchOutputModelList, fileName)
    
    #Excelオブジェクトクラス生成
    createWorkbook = CreateWorkbook.new(xxxGoogleEventList);
#    createWorkbook.setSche(xxxGoogleEventList)
    startDate = Date::new(2013, 1, 20)
    endDate = Date::new(2013, 2, 5)
    createWorkbook.setTerm(startDate, endDate)
  #  createWorkbook.setMember(MemberList)
    #Excel出力実行
    xlbook = createWorkbook.doExe();

    # データ取得。
    
    # データ出力。
    
    
    render :text => 'Hello!'
  end

#  def doExe(serchCondtionModel, fileName)
 #    output(serchCondtionModel, fileName)
  #end

  def getCondi()
    puts "Start"
      #検索条件の設定　Start
      # パスワード
      pass = "パスワード入力"
      # カレンダー情報取得アカウント
      masterAcount = "アカウント入力"
      # アカウントリスト
      acountListMap = {"Aさん"=>"abcdefg@gmail.com", "Bさん"=>"hijklmn@gmail.com", "Cさん"=>"opqrstu@gmail.com", }
      #最大取得件数
      maxResults = 99999
      #開始日
      startMin = "2013-01-01"
      #終了日
      startMax = "2013-12-01"
      #ソート条件
      orderBy = "starttime"
      #ソート順
      sortOrder = "ascending"
      #ファイル名
      fileName = "GoogleCal出力"
      
      serchCondtionModel = SerchCondtionModel.new(pass,masterAcount,acountListMap,maxResults,startMin,startMax,orderBy,sortOrder);
  end
  
  def getXXXXGoogleAPIEventList(serchCondtionModel, fileName)
    #カレンダー情報取得クラスオブジェクト生成
    serchGoogleCalender = SerchGoogleCalender.new(serchCondtionModel);
    #カレンダー情報取得
    serchOutputModelList = serchGoogleCalender.getGoogleCalenderData();

      return serchOutputModelList;
  end
end
