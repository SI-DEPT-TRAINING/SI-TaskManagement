# -*- coding: utf-8 -*-
class ApplicationController < ActionController::Base
  after_filter :convert_to_utf8
  protect_from_forgery :except => :complete

  private
  def convert_to_utf8
    response.charset = "UTF-8"
    response.body = response.body.encode("UTF-8")
  end
end
