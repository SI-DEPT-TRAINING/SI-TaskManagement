module ErrorHandling
  
  # 500エラー
  def render_500(exception = nil)
    if exception
      logger.error "Rendering 500 with exception: #{exception.message}"
    end
    render :file => "#{Rails.root}/public/500.html", :status => 500, :layout => false, :content_type => 'text/html'
  end

  # 404エラー
  def render_404(exception = nil)
    if exception
      logger.error "Rendering 404 with exception: #{exception.message}"
    end
    render :file => "#{Rails.root}/public/404.html", :status => 404, :layout => false, :content_type => 'text/html'
  end
end