# coding: utf-8
require 'searchCondtionModel.rb'
require 'acountModel.rb'
require 'validetionModule.rb'
require "yaml"
require "oauthController.rb"

class GcalSearchController < OAuthController

  @acount = nil
  @password = nil
  @dateFrom = nil
  @dateTo = nil
  @acountList = Array.new;
  @errorMsgList = Array.new;
  @line = nil
  
=begin 
検索画面初期表示
=end
  def index
    @acountList = Array.new;
    @errorMsgList = Array.new;
    render :template => 'gcal_search/index'
  end

=begin 
Ajax通信
=end
  def ajaxSetSession
    setSession
    @respondData = {state: 'ok'}
    respond_to do |format| 
      format.json { render json: @respondData}
    end
  end

=begin 
CSVファイルアップロード
=end
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

=begin 
Excelファイル出力
=end
  def excelOut
    @acountList = Array.new
    @errorMsgList = Array.new

    unless excelValidate then
      
      searchCondtionModel = SearchCondtionModel.new(
                               params[:password],
                               params[:acount],
                               session[:acountList],
                               SYSTEM_YAML["cal_api_maxResults"],
                               params[:dateFrom],
                               params[:dateTo],
                               SYSTEM_YAML["cal_api_orderBy"],
                               SYSTEM_YAML["cal_api_sortOrder"]
                           )
     #googleCalApi
     #ExcelOutModule

      respond_to do |format|
        format.html
        format.xls { output_excel('Excelオブジェクト', "XXX.xls") }
      end
      return
    end

    setSession
    getSession
    render :template => 'gcal_search/index'
  end

=begin 
GoogleOAuth2.0 コールバックアクション
=end
  def oauth2callback
    @acountList = Array.new;
    @errorMsgList = Array.new;
    setGoogleToken
    render :template => 'gcal_search/index'
  end

=begin 
Excel出力時バリデート
=end
  private
  def excelValidate

    isValidate = false
 
    #アカウント名
    checkerAcount = ValidetionModule::MustChecker.new("アカウント", params[:acount])
    if checkerAcount.error then
      @errorMsgList << checkerAcount.errorMsg
      isValidate = true
    end

    #パスワード
    checkerPass = ValidetionModule::MustChecker.new("パスワード", params[:password])
    if checkerPass.error then
      @errorMsgList << checkerPass.errorMsg
      isValidate = true
    end
    
    #日付系
    isValidate = dateValidate

    #CSVデータ
    checkerCsvData = ValidetionModule::MustChecker.new("CSVデータ", session[:acountList])
    if checkerCsvData.error then
      @errorMsgList << "CSVデータを取り込んでください。"
      isValidate = true
    end
    
  return isValidate
  end

=begin 
日付バリデート
=end
  private
  def dateValidate
    isValidate = false

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

=begin 
CSVファイルバリデート
=end
  private
  def csvValidate
    
    csvChecker = ValidetionModule::CsvChecker.new(params[:file])
    if csvChecker.error then
      @errorMsgList << csvChecker.errorMsg
      return true
    end

    csvFile = params[:file]['csv']
    csvChecker.rows(csvFile)
    if csvChecker.error then
      @errorMsgList << csvChecker.errorMsg
      return true
    else
      @line = csvChecker.line
    end

    csvChecker.contentTyoe(csvFile)
    if csvChecker.error then
      @errorMsgList << csvChecker.errorMsg
      return true
    end

    return false
  end

=begin 
POST値をSessionへ保存
=end
  private
  def setSession
    session[:acount] = params[:acount]
    session[:password] = params[:password]
    session[:dateFrom] = params[:dateFrom]
    session[:dateTo] = params[:dateTo]
  end

=begin 
Session値をフォームへ設定
=end
  private
  def getSession
      @acount = session[:acount]
      @password = session[:password]
      @dateFrom = session[:dateFrom]
      @dateTo = session[:dateTo]
      
      if session[:acountList] != nil then
        @acountList = session[:acountList]
      end
  end

end
