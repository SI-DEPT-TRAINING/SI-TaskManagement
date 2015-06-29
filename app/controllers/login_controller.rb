# coding: utf-8
class LoginController < ApplicationController
  after_filter :convert_to_utf8

  # 002
  def index
    reset_session
  end

  # 003
  def googleOAuth
    redirect_to :controller => 'gcalSearch', :action => 'index'
  end

end
