# -*- encoding: utf-8 -*-

class SearchCondtionModel
  def initialize(pass, masterAcount, acountListModel, maxResults, startMin, startMax, orderBy, sortOrder)
    # �p�X���[�h
    @pass = pass;
    # �J�����_�[���擾�A�J�E���g
    @masterAcount = masterAcount;
    # �A�J�E���g���X�g
    @acountListModel = acountListModel;
    #�ő�擾����
    @maxResults = maxResults;
    #�J�n��
    @startMin = startMin;
    #�I����
    @startMax = startMax;
    #�\�[�g����
    @orderBy = orderBy;
    #�\�[�g��
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