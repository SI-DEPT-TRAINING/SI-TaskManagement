# -*- encoding: utf-8 -*-
class SerchCondtionModel
  def initialize
  end
  # �p�X���[�h
  @pass = nil;
  # �J�����_�[���擾�A�J�E���g
  @masterAcount  = nil;
  # �A�J�E���g���X�g
  @acountListMap = nil;
  #�ő�擾����
  @maxResults = nil;
  #�J�n��
  @startMin = nil;
  #�I����
  @startMax = nil;
  #�\�[�g����
  @orderBy = nil;
  #�\�[�g��
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