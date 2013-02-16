# -*- encoding: utf-8 -*-
require "rubygems"
require "google/api_client"
require "yaml"
require "time"

class GcalSessionManager



  @session = nil
  #コンストラクタ
  def initialize(session)
     @session = session
     token_info = session[:gcalToken]
     token_info.blank? ? getGcalToken : checkGcalToken;
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
    token_info = @session[:gcalToken]
    if Time.now > (token_info['issue_timestamp'] + token_info['expires'])
      refresh_GcalToken
    end
  end

  #トークン再発行
  private
  def refresh_GcalToken
      client = Google::APIClient.new
      client.authorization.client_id = CLIENT_ID
      client.authorization.client_secret = CLIENT_SECRET

      #GOOGLE APIが最初に取得したtoken_infoを渡すと情報を初期化してくれるらしい
      client.authorization.update_token!(token_info)#続けてAPI接続しない場合は、"!"を外す??
      client.authorization.grant_type = 'refresh_token'#再発行を意味する
      token_info = client.authorization.fetch.access_token! #Token情報を取得　続けてAPI接続しない場合は、"!"を外す
      token_info['issue_timestamp'] = Time.now
      @session[:gcalToken] = token_info
  end

end