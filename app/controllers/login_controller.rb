class LoginController < ApplicationController
  
  def index
  end

  def googleAuth
    redirect_to :controller => 'gcalSearch', :action => 'index'
  end

end
