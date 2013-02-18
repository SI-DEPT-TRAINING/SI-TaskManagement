# coding: utf-8
class LoginController < ApplicationController
  
  def index
    reset_session
  end

  def googleOAuth
    redirect_to :controller => 'gcalSearch', :action => 'index'
  end

end
