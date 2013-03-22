# coding: utf-8
require 'acountModel.rb'
require 'validetionModule.rb'
require "yaml"
require "oauthController.rb"
require "googleManager.rb"
require "CreateWorkbook.rb"
#require 'win32ole'

class GcalSearchController < OAuthController

  @dateFrom = nil
  @dateTo = nil
  @acountList = Array.new;
  @errorMsgList = Array.new;
  @line = nil
 
  #検索画面初期表示
  def index
    @acountList = Array.new;
    @errorMsgList = Array.new;
    render :template => 'gcal_search/index'
  end
  
  #Ajax通信
  def ajaxSetSession
    setSession
    @respondData = {state: 'ok'}
    respond_to do |format| 
      format.json { render json: @respondData}
    end
  end
  
  #CSVファイルアップロード
  def csvUpLoad

    @acountList = Array.new
    @errorMsgList = Array.new
    unless csvValidate then

      # ----------------------
      # カラム >> 配列
      # ----------------------
      @line.each do |rows|
        acount = AcountModel.new
        row = rows.split(",")
        acount.name = row[0]
        acount.acount = row[1]
        @acountList << acount
      end
    end
    session[:acountList] = @acountList
    getSession
    render :template => 'gcal_search/index'
  end
 
  #Excelファイル出力
  def excelOut

    @acountList = Array.new
    @errorMsgList = Array.new

    unless excelValidate then
 
      #GoogleCalender情報取得
      calResult = getEventList
      #Excelオブジェクト生成
      model = createGcalSearchModel
      excelbook = createWorkbook(calResult, model)
      #Excelダウンロード
      sendExcel(excelbook)

    else

      setSession
      getSession
      render :template => 'gcal_search/index'

    end


  end

  #GoogleOAuth2.0 コールバックアクション
  def oauth2callback

    @acountList = Array.new;
    @errorMsgList = Array.new;
    
    #Token情報をSessionへ保存
    setGoogleToken
    render :template => 'gcal_search/index'
  end
 
  #Excelダウンロード
  private
  def sendExcel(book)
      tmpfile = Tempfile.new ["excel_tmp", ".xls"]
      book.write tmpfile
      tmpfile.open
      filename = "SI-Manage-" + Time.now.strftime('%y%m%d%H%M%S') + ".xls"
      send_data(
        tmpfile.read,
        :disposition=>'attachment',
        :type=>"application/excel",
        :filename => filename
      )
      tmpfile.close(true)
  end

  #GoogleCalender情報取得
  private
  def getEventList
      model = createGcalSearchModel
      calender = GoogleManager::Calender.new(session[:apiClient])
      return calender.getEventList(model)
  end
  
  #Excelオブジェクト生成
  private
  def createWorkbook(calResult, model)
      workbook = CreateWorkbook.new(calResult)
      workbook.setTerm(model.startMin,model.startMax)
      return workbook.doExe
  end
  
  #Excel出力時バリデート
  private
  def excelValidate
    isValidate = false
    #日付系
    isValidate = dateValidate(isValidate)
    #CSVデータ
    checkerCsvData = ValidetionModule::MustChecker.new("CSVデータ", session[:acountList])
    if checkerCsvData.error then
      @errorMsgList << "CSVデータを取り込んでください。"
      isValidate = true
    end
    
  return isValidate
  end

  #日付バリデート
  private
  def dateValidate(_isValidate)
    isValidate = _isValidate

    #日付Form
    checkerDate = ValidetionModule::DateChecker.new("日付Form", params[:dateFrom])
    if checkerDate.error then
      @errorMsgList << checkerDate.errorMsg
      isValidate = true
    end
 
    #日付To
    checkerDate = ValidetionModule::DateChecker.new("日付To", params[:dateTo])
    if checkerDate.error then
      @errorMsgList << checkerDate.errorMsg
      isValidate = true
    end

    #日付共通
    unless isValidate then
      checkerDate.diff(params[:dateFrom], params[:dateTo])
      if checkerDate.error then
        @errorMsgList << checkerDate.errorMsg
        isValidate = true
      end
    end

    return isValidate
  end
 
  #CSVファイルバリデート
  private
  def csvValidate
    
    csvChecker = ValidetionModule::CsvChecker.new(params[:file])
    if csvChecker.error then
      @errorMsgList << csvChecker.errorMsg
      return true
    end

    csvFile = params[:file]['csv']

    csvChecker.contentTyoe(csvFile)
    if csvChecker.error then
      @errorMsgList << csvChecker.errorMsg
      return true
    end

    csvChecker.rows(csvFile)
    if csvChecker.error then
      @errorMsgList << csvChecker.errorMsg
      return true
    else
      @line = csvChecker.line
    end

    return false
  end

  #POST値をSessionへ保存
  private
  def setSession
    session[:dateFrom] = params[:dateFrom]
    session[:dateTo] = params[:dateTo]
  end

  #Session値をフォームへ設定
  private
  def getSession
      @dateFrom = session[:dateFrom]
      @dateTo = session[:dateTo]
      
      if session[:acountList] != nil then
        @acountList = session[:acountList]
      end
  end

  #GoogleCalenderApiモデルクラスを生成します
  private
  def createGcalSearchModel
      model = GoogleManager::CalSearchModel.new
      model.acountListModel = session[:acountList];
      model.startMax = params[:dateTo];
      model.startMin = params[:dateFrom];
      return model;
  end
  
end
