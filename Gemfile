source 'https://rubygems.org'

gem 'rails', '3.2.9'

# Bundle edge Rails instead:
# gem 'rails', :git => 'git://github.com/rails/rails.git'
group :development do
  gem 'sqlite3'
end

# production for Heroku(PostgreSQL)
group :production do
  gem 'pg', '0.15.1'
  gem 'rails_12factor', '0.0.2'

# set JAVA_HOME so Heroku will install gems that need it
heroku_java_home = '/usr/lib/jvm/java-6-openjdk'
ENV['JAVA_HOME'] = heroku_java_home if Dir.exist?(heroku_java_home)

end

# Gems used only for assets and not required
# in production environments by default.
group :assets do
  gem 'sass-rails',   '~> 3.2.3'
  gem 'coffee-rails', '~> 3.2.1'

  # See https://github.com/sstephenson/execjs#readme for more supported runtimes
  # gem 'therubyracer', :platforms => :ruby

  gem 'uglifier', '>= 1.0.3'
end

gem 'jquery-rails'

# To use ActiveModel has_secure_password
# gem 'bcrypt-ruby', '~> 3.0.0'

# To use Jbuilder templates for JSON
# gem 'jbuilder'

# Deploy with Capistrano
# gem 'capistrano'

# To use debugger
# gem 'debugger'

# googleApi
gem 'google-api-client'

# WebServer
#gem 'thin'
gem 'unicorn'

# Excel出力用
gem 'rjb', '1.4.8'

# UnitTest
gem "rspec-rails", "~> 2.0"

# CentOS用
group :production do
  gem 'execjs'
  gem 'therubyracer'
end

# Debugger
group :test, :development do
  # Pry本体
  gem 'pry'
  # デバッカー
  gem 'pry-remote'
  gem 'pry-debugger'
  # show-stackコマンド,framコマンド
  gem 'pry-stack_explorer'
  # Railsコンソールの多機能版
  gem 'pry-rails'
  # PryでのSQLの結果を綺麗に表示
  gem 'hirb'
  gem 'hirb-unicode'
  # pry画面でのドキュメント/ソース表示
  gem "pry-doc"
  # pryの色付けをしてくれる
  gem 'awesome_print'
  # pryの入力に色付け
  gem 'pry-coolline'
end





