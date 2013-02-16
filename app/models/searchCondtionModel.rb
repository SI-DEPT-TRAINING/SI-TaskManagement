# -*- encoding: utf-8 -*-

class SearchCondtionModel
  def initialize(pass, masterAcount, acountListModel, maxResults, startMin, startMax, orderBy, sortOrder)
    # パスワード
    @pass = pass;
    # カレンダー情報取得アカウント
    @masterAcount = masterAcount;
    # アカウントリスト
    @acountListModel = acountListModel;
    #最大取得件数
    @maxResults = maxResults;
    #開始日
    @startMin = startMin;
    #終了日
    @startMax = startMax;
    #ソート条件
    @orderBy = orderBy;
    #ソート順
    @sortOrder = sortOrder;
  end
  attr_accessor :pass
  attr_accessor :masterAcount
  attr_accessor :acountListModel
  attr_accessor :maxResults
  attr_accessor :startMin
  attr_accessor :startMax
  attr_accessor :orderBy
  attr_accessor :sortOrder
end