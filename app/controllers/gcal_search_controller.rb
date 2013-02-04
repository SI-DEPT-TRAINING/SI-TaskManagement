require 'SerchCondtionModel.rb'
require 'AcountInfoModel.rb'
class GcalSearchController < ApplicationController
  @acount = nil;
  @password = nil;
  @dateFrom = nil;
  @dateTo = nil;
  @acountList = nil;
  def index
    @acountList = Array.new;
    render :template => 'gcal_search/index'
  end

  def ajaxSetSession

    session[:acount] = params[:acount]
    session[:password] = params[:password]
    session[:dateFrom] = params[:dateFrom]
    session[:dateTo] = params[:dateTo]
    
    @respondData = {state: 'ok'}
    respond_to do |format| 
      format.json { render json: @respondData}
    end
  end
  def csvUpLoad
    csvFile = params[:file]['csv']
    if csvFile.blank? == false
      
      #------------------------
      # 配列 >>　レコード
      # -----------------------
      line = csvFile.read.split("\n")
      
      # ----------------------
      # カラム >> 配列
      # ----------------------
      @acountList = Array.new;
      line.each do |rows|
        acount = AcountInfoModel.new
        row = rows.split(",")
        acount.name = row[0]
        acount.acount = row[1]
        @acountList << acount
      end
      @acount = session[:acount]
      @password = session[:password]
      @dateFrom = session[:dateFrom]
      @dateTo = session[:dateTo]
      session[:test] = 'acsdv'
    end
    render :template => 'gcal_search/index'
  end
end
