# coding: utf-8
require 'errorHandling.rb'

class ApplicationController < ActionController::Base
  include ErrorHandling
  protected
  def local_request?
    false
  end

  def rescue_action_in_public(exception)
    case exception
    when ActionController::RoutingError,
         ActionController::UnknownAction
        render_404
    else
        render_500
    end
  end

  #フロント：UTF-8対応 test
  # issue_test001 のコメント
  private
  def convert_to_utf8
    response.charset = "UTF-8"
    response.body = response.body.encode("UTF-8")
  end
end
