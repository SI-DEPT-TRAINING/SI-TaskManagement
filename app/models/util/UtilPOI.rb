# -*- encoding: utf-8 -*-
require "rjb";
classpaths = Dir.glob("./lib/*.jar")
Rjb::load(classpath = classpaths.join(','), jvmargs=[])

module JavaIOMod
  FileInputStream = Rjb::import("java.io.FileInputStream")
  FileOutputStream = Rjb::import("java.io.FileOutputStream")
end
module POIMod
  POIFSFileSystem = Rjb::import("org.apache.poi.poifs.filesystem.POIFSFileSystem")
  HSSFWorkbook = Rjb::import("org.apache.poi.hssf.usermodel.HSSFWorkbook")
  HSSFSheet = Rjb::import("org.apache.poi.hssf.usermodel.HSSFSheet")
  HSSFRow = Rjb::import("org.apache.poi.hssf.usermodel.HSSFRow")
  HSSFCell = Rjb::import("org.apache.poi.hssf.usermodel.HSSFCell")
  CellRangeAddress = Rjb::import("org.apache.poi.hssf.util.CellRangeAddress")
#  PoiUtil = Rjb::import("org.bbreak.excella.core.util.PoiUtil")
end

class UtilPOI
  def self.openWorkbook(filename)
    fis = JavaIOMod::FileInputStream.new(filename)
    fs = POIMod::POIFSFileSystem.new(fis);
    book = POIMod::HSSFWorkbook.new(fs)
    fis.close
    return book
  end

  def self.saveWorkbook(book, filename)
    fileOut = JavaIOMod::FileOutputStream.new(filename)
    book.write(fileOut);
    fileOut.close()
  end
  
  def self.deleteCell(sheet, address)
    corner = UtilExcel.getCorner(address)
    ###range = POIMod::CellRangeAddress.new(10, corner[1].to_i, 23, corner[3].to_i)
    range = POIMod::CellRangeAddress.new(1, 1, 20, 20)
    # POIMod::PoiUtil.deleteRangeLeft(sheet, range)
    return ''
  end

  def self.InsertColumn(sheet, colNumber)
    colIndex = colNumber - 1
    for idxRow in 0 .. sheet.getLastRowNum() - 1
      row = sheet.getRow(idxRow)
      if (row == nil)
        next
      end
      currentColIdx = row.getLastCellNum() - 1
      while true
          i = 1
        if (currentColIdx == colIndex)
          oldcell = row.getCell(currentColIdx-i)
          #cell = row.createCell(currentColIdx)
          if oldcell != nil
            style = oldcell.getCellStyle()
            if style != nil
              #cell.setCellStyle(style)
            end
          end
          # break
        end
        if (currentColIdx >= colIndex)
          oldcell = row.getCell(colIndex-i)
          newcell = row.createCell(currentColIdx+i)
          cell = row.getCell(currentColIdx)
          if cell != nil
            # val = cell.getCellValue()
            val = cell.toString()
            if val != nil
              # 右隣のセルへ移動。
              newcell.setCellValue(val)
            end
            if oldcell != nil
              # スタイルをセルへ引き継ぎ。
              style = oldcell.getCellStyle()
              if style != nil
                # cell.setCellStyle(style)
                newcell.setCellStyle(style)
              end
            end
            w = sheet.getColumnWidth(colIndex - 1)
            if w != nil
              sheet.setColumnWidth(currentColIdx, w)
            end
          end
          
          #row.moveCell(cell, currentColIdx+i)
        else
          oldcell = row.getCell(colIndex-i)
          cell = row.createCell(colIndex)
          if oldcell != nil
            style = oldcell.getCellStyle()
            if style != nil
              cell.setCellStyle(style)
            end
          end
          break
        end
        currentColIdx = currentColIdx - 1 
      end
    end

    return nil
  end
  
  def self.getCell(sheet, address)
    corner = UtilExcel.getCorner(address)
    rowIndex = corner[1].to_i - 1
    colIndex = UtilExcel.getColumnNumber(corner[0]) - 1
    row = sheet.getRow(rowIndex)
    if (row == nil)
      row = sheet.createRow(rowIndex)
    end
    cell = row.getCell(colIndex)
    if (cell == nil)
      cell = row.createCell(colIndex)
    end
    
    return cell
  end
  
  def self.getAddress(cell, offsetRow=0, offsetCol=0)
    rowIndex = cell.getRowIndex() + offsetRow
    colIndex = cell.getColumnIndex() + offsetCol
    
    address = UtilExcel.getColumnAlphabet(colIndex + 1) + (rowIndex + 1).to_s
    return address
  end
  
  def self.offset(cell, rowOffset, colOffset)
    sheet = cell.getSheet()
    
    rowIndex = cell.getRowIndex() + rowOffset
    colIndex = cell.getColumnIndex() + colOffset
    row = sheet.getRow(rowIndex)
    if (row == nil)
      row = sheet.createRow(rowIndex)
    end
    cell = row.getCell(colIndex)
    if (cell == nil)
      cell = row.createCell(colIndex)
    end
    
    return cell
  end
  
  def self.setValue(cell, value, offsetRow=0, offsetCol=0)
    if (cell == nil)
      return 
    end
    
    targetCell = cell
    if (offsetRow !=0 || offsetCol != 0 )
      # offset指定があればそのセルを求める。
      targetCell = self.offset(cell, offsetRow, offsetCol)
    end
    # 対象セルに値をセット。
    targetCell.setCellValue(value)
    return 
  end
  
  def self.getValue(cell, offsetRow=0, offsetCol=0)
    if (cell == nil)
      return ""
    end
    
    targetCell = cell
    if (offsetRow !=0 || offsetCol != 0 )
      # offset指定があればそのセルを求める。
      targetCell = self.offset(cell, offsetRow, offsetCol)
    end
    # 対象セルの値を取得。
    return targetCell.toString()
  end
  
  def self.setFormula(cell, ｆormula, offsetRow=0, offsetCol=0)
    if (cell == nil)
      return 
    end
    
    targetCell = cell
    if (offsetRow !=0 || offsetCol != 0 )
      # offset指定があればそのセルを求める。
      targetCell = self.offset(cell, offsetRow, offsetCol)
    end
    # 対象セルに値をセット。
    targetCell.setCellFormula(ｆormula)
    return 
  end
end
