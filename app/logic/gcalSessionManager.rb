# -*- encoding: utf-8 -*-
require "rubygems"
require "google/api_client"
require "yaml"
require "time"

class GcalSessionManager



  @session = nil
  #�R���X�g���N�^
  def initialize(session)
     @session = session
     token_info = session[:gcalToken]
     token_info.blank? ? getGcalToken : checkGcalToken;
  end

  #�g�[�N�����s
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

  #�g�[�N�������؂�`�F�b�N
  private
  def checkGcalToken
    token_info = @session[:gcalToken]
    if Time.now > (token_info['issue_timestamp'] + token_info['expires'])
      refresh_GcalToken
    end
  end

  #�g�[�N���Ĕ��s
  private
  def refresh_GcalToken
      client = Google::APIClient.new
      client.authorization.client_id = CLIENT_ID
      client.authorization.client_secret = CLIENT_SECRET

      #GOOGLE API���ŏ��Ɏ擾����token_info��n���Ə������������Ă����炵��
      client.authorization.update_token!(token_info)#������API�ڑ����Ȃ��ꍇ�́A"!"���O��??
      client.authorization.grant_type = 'refresh_token'#�Ĕ��s���Ӗ�����
      token_info = client.authorization.fetch.access_token! #Token�����擾�@������API�ڑ����Ȃ��ꍇ�́A"!"���O��
      token_info['issue_timestamp'] = Time.now
      @session[:gcalToken] = token_info
  end

end