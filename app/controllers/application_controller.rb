# coding: utf-8
require 'errorHandling.rb'
class ApplicationController < ActionController::Base
  after_filter :convert_to_utf8
  include ErrorHandling
  #protect_from_forgery


  # Exceptionが発生した場合は、一律共通エラー画面へ遷移させる
  # 本アクションが実行されるのは、production で、かつローカル
  # らのリクエストでない場合のみのため、動作確認用に下記設定が必須
  # config/environments/development.rb
  # config.action_controller.consider_all_requests_local = false
  # local_request?のオーバーライド且つfalse
  private
  def rescue_action_in_public(exception)
    case exception
    when ::ActionController::RoutingError
        render_404
    else
        render_500
    end
  end

  private
  def local_request?
      false
  end

  #excel出力
  protected
  def output_excel(template, fname = nil)
    options = fname ? { filename:IEpatch(fname) } : {}
    send_data(render_to_string(template:template, layout:'excel.html'), options)
  end

  #Internet Explorerだとファイル名が文字化けする??
  private
  def IEpatch(fname)
    /MSIE/ =~ request.env['HTTP_USER_AGENT'] ? CGI.escape(fname) : fname
  end

  #フロント：UTF-8対応
  private
  def convert_to_utf8
    response.charset = "UTF-8"
    response.body = response.body.encode("UTF-8")
  end
end
