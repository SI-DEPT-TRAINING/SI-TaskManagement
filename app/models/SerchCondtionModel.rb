# -*- encoding: utf-8 -*-
class SerchCondtionModel
  def initialize
  end
  # パスワード
  @pass = nil;
  # カレンダー情報取得アカウント
  @masterAcount  = nil;
  # アカウントリスト
  @acountListMap = nil;
  #最大取得件数
  @maxResults = nil;
  #開始日
  @startMin = nil;
  #終了日
  @startMax = nil;
  #ソート条件
  @orderBy = nil;
  #ソート順
  @sortOrder = nil;
  attr_accessor :pass
  attr_accessor :masterAcount
  attr_accessor :acountListMap
  attr_accessor :maxResults
  attr_accessor :startMin
  attr_accessor :startMax
  attr_accessor :orderBy
  attr_accessor :sortOrder
end