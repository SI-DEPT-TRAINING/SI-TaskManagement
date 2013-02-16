# coding: utf-8
class LoginController < ApplicationController
  
  def index
    session[:gcalToken] = nil
  end

  def googleOAuth
    redirect_to :controller => 'gcalSearch', :action => 'index'
  end

end
