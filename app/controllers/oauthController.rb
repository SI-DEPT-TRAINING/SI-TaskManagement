# coding: utf-8
require "rubygems"
require "google/api_client"
require "yaml"
require "time"

class OAuthController < ApplicationController
  before_filter :creatToken
  
  #各種パラメータは定義ファイルより取得する
  SYSTEM_YAML = YAML.load_file(File.dirname(__FILE__) + '/../../config/googleSystemProp.yml')
  CLIENT_ID = SYSTEM_YAML["client_id"]
  CLIENT_SECRET = SYSTEM_YAML["client_secret"]
  OUTH_SCOPE = SYSTEM_YAML["scope"]
  REDIRECT_URI = SYSTEM_YAML["redirect_uri"]
  DISCOVER = SYSTEM_YAML["discovered_api"]
  VERSION =  SYSTEM_YAML["verision"]

  #トークン生成
  def creatToken
    
    #コールバックチェック
    unless params[:code].blank?
      return
    end

    #Token生成またTokenリフレッシュ
    token_info = session[:gcalToken]
    token_info.blank? ? getGcalToken : doExpirarion;
  end

  #コールバックアクション
  protected
  def setGoogleToken

    client = Google::APIClient.new
    client.authorization.client_id = CLIENT_ID
    client.authorization.client_secret = CLIENT_SECRET
    client.authorization.redirect_uri = REDIRECT_URI
    client.authorization.code = params[:code] #許可コード取得
    token_info = client.authorization.fetch_access_token! #Token発行
    token_info['origin_timestamp'] = Time.now #期限切れ基準時刻 
    session[:gcalToken] = token_info
    session[:apiClient] = client
  end

  #トークン発行
  private
  def getGcalToken
    client = Google::APIClient.new
    client.authorization.client_id = CLIENT_ID
    client.authorization.client_secret = CLIENT_SECRET
    client.authorization.scope = OUTH_SCOPE
    client.authorization.redirect_uri = REDIRECT_URI
    uri = client.authorization.authorization_uri
    
    begin
      redirect_to uri.to_s
    rescue => exception
      logger.error "OauthApi ERROR: #{exception.message}"
    end

  end
  
  #トークン期限切れチェック
  private
  def doExpirarion
    token_info = session[:gcalToken]
    if Time.now > (token_info['origin_timestamp'] + token_info['expires_in'])
      refreshGcalToken
    end
  end

  #トークン再発行
  private
  def refreshGcalToken

      client = Google::APIClient.new
      client.authorization.client_id = CLIENT_ID
      client.authorization.client_secret = CLIENT_SECRET

      #GOOGLE APIが最初に取得したtoken_infoを渡すと情報を初期化してくれるらしい
      token_info = session[:gcalToken]
      client.authorization.update_token!(token_info)

      client.authorization.grant_type = 'refresh_token' #再発行キー
      token_info = client.authorization.fetch_access_token #Token再発行
      token_info['origin_timestamp'] = Time.now #期限切れ基準時刻
      session[:gcalToken] = token_info
      session[:apiClient] = client
  end
  
end
