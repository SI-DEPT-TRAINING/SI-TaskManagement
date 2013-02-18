# -*- encoding: utf-8 -*-
# require "gcalapi";

class UtilExcel
  # 
  def self.getColumnNumber(columnAlpha)
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
  
  def self.getCorner(cellAddress)
    /(\$?([A-Z]+)\$?([0-9]+))(:(\$?([A-Z]+)\$?([0-9]+)))?/ =~ cellAddress
    if $5 == nil
      return [ $2, $3, $2, $3 ]
    end
    return [ $2, $3, $6, $7 ]
  end
  def self.getRowsAddress(cellAddress)
    array = self.getCorner(cellAddress)
    rows_address = sprintf("%s:%s", array[1], array[3])
    return rows_address
  end
  def self.getColumnsAddress(cellAddress)
    array = self.getCorner(cellAddress)
    columns_address = sprintf("%s:%s", array[0], array[2])
    return columns_address
  end
  def self.getTopLeftAddress(cellAddress)
    array = self.getCorner(cellAddress)
    top_left_address = sprintf("%s%s", array[0], array[1])
    return top_left_address
  end
  def self.getBottomRightAddress(cellAddress)
    array = self.getCorner(cellAddress)
    bottom_right_address = sprintf("%s%s", array[2], array[3])
    return bottom_right_address
  end
end
