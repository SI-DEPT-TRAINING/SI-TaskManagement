# coding: utf-8
require "rubygems"
require "google/api_client"
require "yaml"
require "time"

class OAuthController < ApplicationController
  before_filter :creatToken
  
  #各種パラメータは定義ファイルより取得する
  SYSTEM_YAML = YAML.load_file(File.dirname(__FILE__) + '/../../config/systemprop.yml')
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

    token_info = session[:gcalToken]
    token_info.blank? ? getGcalToken : checkGcalToken;
  end

  #コールバックアクション
  protected
  def setGoogleToken

    client = Google::APIClient.new
    client.authorization.client_id = CLIENT_ID
    client.authorization.client_secret = CLIENT_SECRET
    client.authorization.redirect_uri = REDIRECT_URI
    client.authorization.code = params[:code] #許可コード取得
    token_info = client.authorization.fetch_access_token! #続けてAPI接続しない場合は、"!"付与
    token_info['issue_timestamp'] = Time.now

    session[:gcalToken] = token_info
  end

  #トークン発行
  private
  def getGcalToken
    client = Google::APIClient.new
    service = client.discovered_api('calendar', 'v3')
    client.authorization.client_id = CLIENT_ID
    client.authorization.client_secret = CLIENT_SECRET
    client.authorization.scope = OUTH_SCOPE
    client.authorization.redirect_uri = REDIRECT_URI
    uri = client.authorization.authorization_uri
    redirect_to uri.to_s
  end
  
  #トークン期限切れチェック
  private
  def checkGcalToken
    token_info = session[:gcalToken]
    if Time.now > (token_info['issue_timestamp'] + token_info['expires_in'])
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
      client.authorization.update_token!(token_info)#続けてAPI接続しない場合は、"!"を外す??
      client.authorization.grant_type = 'refresh_token'#再発行を意味する
      token_info = client.authorization.fetch_access_token! #Token情報を取得　続けてAPI接続しない場合は、"!"を付与
      token_info['issue_timestamp'] = Time.now
      session[:gcalToken] = token_info
  end
  
end
