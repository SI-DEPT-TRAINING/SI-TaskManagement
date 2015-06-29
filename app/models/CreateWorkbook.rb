# -*- coding: utf-8 -*-
require './app/models/ScheduleList.rb'
require './app/models/ScheduleModel.rb'
require './app/models/util/UtilPOI.rb'
require './app/models/util/UtilExcel.rb'
require './app/models/util/UtilFile.rb'
module Excel
end

classpaths = Dir.glob("./lib/*.jar")
#Rjb::load(classpath = classpaths.join(';'), jvmargs=[]) kogane

module POIMod
  POIFSFileSystem = Rjb::import("org.apache.poi.poifs.filesystem.POIFSFileSystem")
  HSSFWorkbook = Rjb::import("org.apache.poi.hssf.usermodel.HSSFWorkbook")
  HSSFSheet = Rjb::import("org.apache.poi.hssf.usermodel.HSSFSheet")
  HSSFRow = Rjb::import("org.apache.poi.hssf.usermodel.HSSFRow")
  HSSFCell = Rjb::import("org.apache.poi.hssf.usermodel.HSSFCell")
end

class CreateWorkbook

  # コンストラクタ。
  def initialize(googleEventList)
    @googleEventList = googleEventList

    # 引数のスケジュールからメンバーを取得する。
    @memberList = Array.new()
    googleEventList.each do |serchOutputModel|
      member_name = serchOutputModel.name
      @memberList.push(member_name)
    end
  end

  def setTerm(startDate, endDate)
    @startDate = startDate
    @endDate = endDate
  end

  def setMemberList(memberList)
    @memberList = memberList
  end

  def doExe
    system = Rjb::import("java.lang.System")
    # データ取得。
    scheduleList = getExcelScheduleList()
    # データ書き込み。
    excelbook = getExcelWorkbook(scheduleList);
    return excelbook;
  end

  def getExcelScheduleList()
    #デバック用（本来は画面表示用のオブジェクトへ詰め替える予定）
    scheduleList = ScheduleList.new()
    @googleEventList.each do |serchOutputModel|
      if (serchOutputModel.eventList == nil)
        return scheduleList;
      end
      serchOutputModel.eventList.each do |event|
        schedule = ScheduleModel.new(event);
        schedule.username = serchOutputModel.name
        scheduleList.add(schedule)
#        print(cnt , "件目".encode("Shift_JIS"))
      end
    end
    return scheduleList;
  end
  private :getExcelScheduleList

  def outputDateTerm(sheet)
    # セルの行列削除。
    delele_col = UtilExcel.getColumnsAddress(@setting["InitDataCellRange"])
    delele_row = UtilExcel.getRowsAddress(@setting["InitDataCellRange"])
    delete_cell = @setting["InitDataCellRange"]
#    UtilPOI.deleteCell(sheet, delete_cell)
    ####sheet.Range(delele_col).Delete
    ####sheet.Range(delele_row).Delete

    # 日付列の開始位置を取得。
    date_start_address = @setting["BaseCell_Date"]
    date_start_col = UtilExcel.getCorner(date_start_address)[0]

    # 必要分の日付列を一括で列Insert。
    term = (@endDate - @startDate).to_i + 1
    date_end_col = UtilExcel.getColumnAlphabet_onBase(date_start_col, term - 1)
    ### sheet.Range(date_start_col + ":" + date_end_col).Insert()

    for curr in @startDate..@endDate
      offset = (curr - @startDate).to_i
      UtilPOI.InsertColumn(sheet, 11+offset)
      # 日付入力
      row = sheet.getRow(7)
      if (row == nil)
        row = sheet.createRow(7)
      end
      cell = row.getCell(10+offset)
      if (cell == nil)
        cell = row.createCell(10+offset)
      end
      if (cell != nil)
        str = sprintf("%d月%d日", curr.month, curr.day)
        cell.setCellValue(str)
      end

      # 曜日入力
      wdays = ["日", "月", "火", "水", "木", "金", "土"]
      row = sheet.getRow(8)
      if (row == nil)
        row = sheet.createRow(8)
      end
      cell = row.getCell(10+offset)
      if (cell == nil)
        cell = row.createCell(10+offset)
      end
      if (cell != nil)
        str = wdays[curr.wday]
        cell.setCellValue(str)
      end
    end

#    base_cell = sheet.Range(date_start_address)
 #   for curr in @startDate..@endDate
  #    offset = (curr - @startDate).to_i
   #   # 日付、曜日の入力。
    #  base_cell.Offset(0, offset).Value = curr.to_s
     # base_cell.Offset(1, offset).Formula = '=' + base_cell.Offset(0, offset).Address() 
    #end
  end
  private :outputDateTerm

  def outputSchedule(sheet, scheduleList)
    addIdx = 0
    errIdx = 0
    error_base_cell_str = @setting["BaseCell_Error"]
    work_time_unit = @setting["WorkTimeUnit"].to_i
    time_type = @setting["TimeType"]
    address = @setting["BaseCell_ScheduleItem"]
    scheduleList.getList.each do |schedule|
      # スケジュールデータか確認する。
      flg = schedule.isProcessTarget?()
      if (flg == true)
        # スケジュールデータ。
        ####base_cell = sheet.Range(address).Offset(addIdx, 0)
        cell = UtilPOI.getCell(sheet, address)
        base_cell = UtilPOI.offset(cell, addIdx, 0)
        # 分類のリストを取得。
        sectionList = schedule.getSectionList()
        idx = schedule.getSameSectionIndexBeforeMyself(scheduleList)
        isNewRow = false
        if (idx < 0)
          isNewRow = true
        end
        if (idx < 0)
          # 同一分類のスケジュールが存在しない。
          ins_ = base_cell.getRowIndex()
          sheet.shiftRows(ins_, sheet.getLastRowNum(), 1)
          row = sheet.getRow(ins_)
          if (row == nil)
            sheet.createRow(ins_)
          end
          base_cell = UtilPOI.offset(base_cell, -1, 0)
          ins_ = base_cell.getRowIndex()

          UtilPOI.setValue(base_cell, getNendoYY(schedule.getStartDate()), 0, 0)
          UtilPOI.setValue(base_cell, sectionList[0], 0, 1)
          UtilPOI.setValue(base_cell, sectionList[1], 0, 2)
          UtilPOI.setValue(base_cell, sectionList[2], 0, 3)

          tmpAdd = UtilPOI.getAddress(base_cell, 0, 1)
          fomular = "VLOOKUP(" + tmpAdd + ", コード定義!$C$4:$D$100, 2, FALSE)"
          UtilPOI.setFormula(base_cell, fomular, 0, 4)

          tmpAdd = UtilPOI.getAddress(base_cell, 0, 2)
          fomular = "VLOOKUP(" + tmpAdd + ", コード定義!$F$4:$G$100, 2, FALSE)"
          UtilPOI.setFormula(base_cell, fomular, 0, 5)

          tmpAdd = UtilPOI.getAddress(base_cell, 0, 3)
          fomular = "VLOOKUP(" + tmpAdd + ", コード定義!$I$4:$J$100, 2, FALSE)"
          UtilPOI.setFormula(base_cell, fomular, 0, 6)

          # 縦軸（分類）ごとの合計式を設定。
          offset1 = (@endDate - @startDate) + 9
          offset1 = offset1.to_i
          /(\$?([A-Z]+)\$?([0-9]+))(:(\$?([A-Z]+)\$?([0-9]+)))?/ =~ @setting["InitDataCellRange"]
          address_col1 = $2
      ####    base_cell.Offset(0, offset1).Formula = "=SUM(" + address_col1 + ins_ + ":INDIRECT(ADDRESS(ROW(), COLUMN()-1)))"
          ### fomular = "SUM(" + address_col1 + (ins_ + 1).to_s + ":INDIRECT(ADDRESS(ROW(), COLUMN()-1)))"
          tmpAdd = UtilPOI.getAddress(base_cell, 0, offset1-1)
          fomular = "SUM(K" + (ins_ + 1).to_s + ":" + tmpAdd + ")"
          UtilPOI.setFormula(base_cell, fomular, 0, offset1)
        else
          # 同一分類のスケジュールが存在する。
          ###base_cell = sheet.Range(address).Offset(idx, 0)
          margedIndex = scheduleList.getMargeSectionIndex(schedule)
          cell = UtilPOI.getCell(sheet, address)
          base_cell = UtilPOI.offset(cell, margedIndex, 0)
        end

        scheDateStr = schedule.getStartDate()
        scheDate = Date::new(scheDateStr.year, scheDateStr.month, scheDateStr.day)
        colOffset = scheDate - @startDate
        # if文はいらないと思うが念のため。出力期間外のスケジュールは書き込まない配慮。
        if (0 <= colOffset)
          if (colOffset <= (@endDate - @startDate))
            # 出力期間内のスケジュール。書き込み。
            val1 = UtilPOI.getValue(base_cell, 0, colOffset.to_i + 8)
            if (val1.blank?)
              val1 = 0
            end
            # 表示形式（h：時間、h以外：分）
            if (time_type == "h")
              val2 = schedule.getWorkTimeHours(work_time_unit)
            else
              val2 = schedule.getWorkTimeMinuts(work_time_unit)
            end
            UtilPOI.setValue(base_cell, val1.to_f + val2.to_f, 0, colOffset.to_i + 8)
          end
        end

        if (idx < 0)
          addIdx += 1
        end
      else
        # エラーデータ。
        cell = UtilPOI.getCell(sheet, error_base_cell_str)
        base_cell = UtilPOI.offset(cell, addIdx + errIdx, 0)
        # エラー用セルに期間とタイトルを出力。
        UtilPOI.setValue(base_cell, schedule.getTermString())
        UtilPOI.setValue(base_cell, schedule.getTitle(), 0, 5)
        errIdx += 1
      end
    end

    if addIdx == 0
      return ;
    end
    # 合計列。
    # 横軸（日付）ごとの合計式を設定。
    work_time_base_address = @setting["BaseCell_WorkTime"]
    /(\$?([A-Z]+)\$?([0-9]+))(:(\$?([A-Z]+)\$?([0-9]+)))?/ =~ work_time_base_address
    work_time_base_address_row = $3

    cell = UtilPOI.getCell(sheet, work_time_base_address)
    base_cell = UtilPOI.offset(cell, addIdx, 0)
    for curr_date in @startDate..@endDate+1
      offset = (curr_date - @startDate).to_i

      curr_cell = UtilPOI.offset(base_cell, 0, offset)
      address_col1 = UtilExcel.getColumnAlphabet(curr_cell.getColumnIndex() + 1)
      address1 = sprintf("%s%s", address_col1, work_time_base_address_row)
      address2 = UtilPOI.getAddress(curr_cell, -1, 0)
      ####base_cell.Offset(0, offset).Formula = "=SUM(" + address + ":INDIRECT(ADDRESS(ROW()-1, COLUMN())))"
      formula = "SUM(" + address1 + ":" + address2 + ")"
      UtilPOI.setFormula(base_cell, formula, 0, offset)
    end
  end

  def getExcelObj()
    begin
      # 既存のExcelプロセスが起動していれば、それを使いまわす。
      excelapp = WIN32OLE::connect("Excel.Application")
    rescue WIN32OLERuntimeError
      # Excelプロセスがないので、新たにプロセスを作成。
      excelapp = WIN32OLE.new("Excel.Application")
    end

    # WIN32OLE.const_load(excelapp, Excel)
    return excelapp
  end

  def getTemplateFileFullpath()
    template_base_path = 'app/assets/excel/'
    template_abspath = template_base_path + '/template.xls'
    return template_abspath
  end

  def getExcelWorkbook(scheduleList)
    # Excelのプロセスを取得する。
    # @excelapp = getExcelObj()

    # テンプレートからファイル作成。
    template_fullpath = getTemplateFileFullpath()

    # TODO
    #time_string = Time.now.strftime("%Y%m%d_%H%M%S")
    #dest = UtilFile.getParentPath(template_fullpath) + '/output' + time_string + '.xls'
    #FileUtils.copy(template_fullpath, dest)

    #book = UtilPOI.openWorkbook(dest)
    book = UtilPOI.openWorkbook(template_fullpath)

    # 「設定」シートから設定の読み込み。
    loadSetting(book)

    @memberList.each do |outputMember|
      # 対象メンバーのスケジュールのみを取得。
      memberScheduleList = scheduleList.narrowMember(outputMember)
      # シート作成。
      sheet = createMemberSheet(book, outputMember)
      # 日付列出力。
      outputDateTerm(sheet)
      # シートにスケジュール出力。
      outputSchedule(sheet, memberScheduleList)

      # 列幅調整。
      #autoFillColumnWidth(sheet)
    end

    # 全メンバーのスケジュールのみを取得。
    # シート作成。
    sheet = createMemberSheet(book, 'ALL')
    # 日付列出力。
    outputDateTerm(sheet)
    # シートにスケジュール出力。
    outputSchedule(sheet, scheduleList)
    # 列幅調整。
    #autoFillColumnWidth(sheet)

    return book;
  end
  private :getExcelWorkbook

  def loadSetting(book)
    @setting = {}
    #sheet = book.worksheets("設定")
    sheet = book.getSheet("設定")
    i = 3
    while true do
      # curr = sheet.Range("C" + i.to_s)
      row = sheet.getRow(i)
      if (row == nil)
        # 最終行。
        break
      end
      curr = row.getCell(2)

      #if (curr.value == nil || curr.value == "" )
      name = curr.toString()
      if (name == nil || name == "" )
        # 設定名称がないならbreak。
        break
      end

      # value = curr.Offset(0, 1).getStringCellValue()
      value = row.getCell(3).toString()
      @setting[name] = value
      i += 1
    end
  end
  private :loadSetting

  def autoFillColumnWidth(sheet)
    date_start_address = @setting["BaseCell_Date"]
    base_cell = UtilPOI.getCell(sheet, date_start_address)
    for curr_date in @startDate..@endDate+1
      offset = (curr_date - @startDate).to_i
      # 日付セル位置の取得
      cell = UtilPOI.offset(base_cell, 0, offset)
      cellIndex = cell.getColumnIndex()
      # 日付列の開始位置を取得。
      sheet.autoSizeColumn(cellIndex)
    end
  end

  def createMemberSheet(book, outputMember)
    # テンプレート用シートのインデックスを取得。
    sheetIdx = book.getSheetIndex('templateSheet')
    # テンプレート用シートからシートの複製。
    cloneSheet = book.cloneSheet(sheetIdx)
    # 複製したシートの位置を変えるとセル式のバグあり。
    # どうにもならなそうなのでシート位置は入れ替えないこととする。

    # コピー先のシートインデックスを取得。
    sheetName = cloneSheet.getSheetName()
    sheetIdx = book.getSheetIndex(sheetName)
    # コピー先のシート名を変更
    book.setSheetName(sheetIdx, outputMember)

    return cloneSheet
  end
  private :createMemberSheet

  def getNendoYY(date)
    tmp = date.year
    if (date.month < 4 )
      tmp -= 1
    end
    return tmp.modulo(100)
  end
end
