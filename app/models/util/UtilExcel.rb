# -*- encoding: utf-8 -*-
# require "gcalapi";

class UtilExcel
  # 
  def self.getColumnNumber(columnAlpha)
    if (/^[A-Za-z]+$/ =~ columnAlpha) then
      # 大文字小文字を無視するため、引数を大文字化する。
      columnAlpha = columnAlpha.upcase
    else
      # 引数がマッチしない。
      return -1
    end
    i = 0
    total = 0
    alphaList = columnAlpha.chars.to_a
    alphaList.reverse.each do |alpha|
      idx = ('A'..'Z').to_a.index(alpha) + 1
      tmp = idx * (26 ** i)
      total += tmp
      i = i + 1
    end
    
    return total
  end
  # 
  def self.getColumnAlphabet_onBase(baseAlpha, colNumber)
    tmp = getColumnNumber(baseAlpha)
    alpha_str = self.getColumnAlphabet(tmp + colNumber)
    return alpha_str
  end
  def self.getColumnAlphabet(colNumber)
    if colNumber <= 0
      return nil
    end
    
    # 下位の位からアルファベット化していく。
    alpha_str = ""
    tmp = colNumber - 1
    while (true)
      ascii_int = 65 + (tmp % 26).to_i
      alpha = ascii_int.chr
      alpha_str = alpha.to_s + alpha_str
      tmp = tmp / 26
      tmp = tmp.to_i - 1
      if tmp < 0
        break
      end
    end
    return alpha_str 
  end
  
  def self.isCellAddress?(cellAddress)
    /^(\$?([A-Z]+)\$?([0-9]+))(:)?((\$?([A-Z]+)\$?([0-9]+)))?$/ =~ cellAddress
    if $4 == ":"
      if (/^(\$?([A-Z]+)\$?([0-9]+))?/ =~ $5) then
          return true
      end
      return false
    else
      if (/^(\$?([A-Z]+)\$?([0-9]+))?/ =~ $1) then
          return true
      end
      return false
    end
    
    return nil
  end
  def self.getCorner(cellAddress)
    if isCellAddress?(cellAddress) == false then
      return nil
    end
    
    # /(\$?([A-Z]+)\$?([0-9]+))(:(\$?([A-Z]+)\$?([0-9]+)))?/ =~ cellAddress
    if (/(\$?([A-Z]+)\$?([0-9]+))(:)?((\$?([A-Z]+)\$?([0-9]+)))?/ =~ cellAddress) then
    else
      return nil
    end
    
    if $4 == nil
      return [ $2, $3, $2, $3 ]
    end
    if $5 == nil
      return nil
    end
    return [ $2, $3, $7, $8 ]
  end
  def self.getRowsAddress(cellAddress)
    if isCellAddress?(cellAddress) == false then
      return nil
    end
    array = self.getCorner(cellAddress)
    rows_address = sprintf("%s:%s", array[1], array[3])
    return rows_address
  end
  def self.getColumnsAddress(cellAddress)
    if isCellAddress?(cellAddress) == false then
      return nil
    end
    array = self.getCorner(cellAddress)
    columns_address = sprintf("%s:%s", array[0], array[2])
    return columns_address
  end
  def self.getTopLeftAddress(cellAddress)
    if isCellAddress?(cellAddress) == false then
      return nil
    end
    array = self.getCorner(cellAddress)
    top_left_address = sprintf("%s%s", array[0], array[1])
    return top_left_address
  end
  def self.getBottomRightAddress(cellAddress)
    if isCellAddress?(cellAddress) == false then
      return nil
    end
    array = self.getCorner(cellAddress)
    bottom_right_address = sprintf("%s%s", array[2], array[3])
    return bottom_right_address
  end
end
