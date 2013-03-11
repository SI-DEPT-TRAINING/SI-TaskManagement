# -*- coding: utf-8 -*-
require 'win32ole'
require './app/models/util/UtilExcel.rb'
module Excel
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
    sheet.Range(delele_col).Delete
    sheet.Range(delele_row).Delete
    
    # 日付列の開始位置を取得。
    date_start_address = @setting["BaseCell_Date"]
    date_start_col = UtilExcel.getCorner(date_start_address)[0]
    
    # 必要分の日付列を一括で列Insert。    
    term = (@endDate - @startDate).to_i + 1
    date_end_col = UtilExcel.getColumnAlphabet_onBase(date_start_col, term - 1)
    sheet.Range(date_start_col + ":" + date_end_col).Insert()
    
    base_cell = sheet.Range(date_start_address)
    for curr in @startDate..@endDate
      offset = (curr - @startDate).to_i
      # 日付、曜日の入力。
      base_cell.Offset(0, offset).Value = curr.to_s
      base_cell.Offset(1, offset).Formula = '=' + base_cell.Offset(0, offset).Address() 
    end
  end
  private :outputDateTerm
  
  def outputSchedule(sheet, scheduleList)
    addIdx = 0
    errIdx = 0
    error_base_cell_str = @setting["BaseCell_Error"]
    work_time_unit = @setting["WorkTimeUnit"].to_i
    time_type = @setting["TimeType"]
    address = @setting["BaseCell_ScheduleItem"]
    scheduleList.each do |schedule|
      # スケジュールデータか確認する。
      flg = schedule.isProcessTarget?()
      if (flg == true) 
        # スケジュールデータ。
        base_cell = sheet.Range(address).Offset(addIdx, 0)
        # 分類のリストを取得。
        sectionList = schedule.getSectionList()
        idx = schedule.getSameSectionIndexBeforeMyself(scheduleList)
        isNewRow = false
        if (idx < 0)
          isNewRow = true
        end
        if (idx < 0)
          # 同一分類のスケジュールが存在しない。
          ins_ = base_cell.Row().to_s
          sheet.Range(ins_ + ":" + ins_).Insert()
          base_cell = base_cell.Offset(-1, 0)
          
          # sheet.Range("A1:C1").Value = [ ["a", "V", "c"] ]
          base_cell.Offset(0, 0).Value = getNendoYY(schedule.getStartDate())
          base_cell.Offset(0, 1).Value = sectionList[0]
          base_cell.Offset(0, 2).Value = sectionList[1]
          base_cell.Offset(0, 3).Value = sectionList[2]
          base_cell.Offset(0, 4).Formula = "=VLOOKUP(" + base_cell.Offset(0, 1).Address() + ", コード定義!$C$4:$D$100, 2, FALSE)"
          base_cell.Offset(0, 5).Formula = "=VLOOKUP(" + base_cell.Offset(0, 2).Address() + ", コード定義!$F$4:$G$100, 2, FALSE)"
          base_cell.Offset(0, 6).Formula = "=VLOOKUP(" + base_cell.Offset(0, 3).Address() + ", コード定義!$I$4:$J$100, 2, FALSE)"
          # 縦軸（分類）ごとの合計式を設定。
          offset1 = (@endDate - @startDate) + 9
          offset1 = offset1.to_i
          /(\$?([A-Z]+)\$?([0-9]+))(:(\$?([A-Z]+)\$?([0-9]+)))?/ =~ @setting["InitDataCellRange"]
          address_col1 = $2
          base_cell.Offset(0, offset1).Formula = "=SUM(" + address_col1 + ins_ + ":INDIRECT(ADDRESS(ROW(), COLUMN()-1)))"
        else
          # 同一分類のスケジュールが存在する。
          base_cell = sheet.Range(address).Offset(idx, 0)
        end

        scheDateStr = schedule.getStartDate()
        scheDate = Date::new(scheDateStr.year, scheDateStr.month, scheDateStr.day)
        colOffset = scheDate - @startDate
        # if文はいらないと思うが念のため。出力期間外のスケジュールは書き込まない配慮。
        if (0 <= colOffset)
          if (colOffset <= (@endDate - @startDate))
            # 出力期間内のスケジュール。書き込み。
            val1 = base_cell.Offset(0, colOffset.to_i + 8).Value
            if (val1.blank?)
              val1 = 0
            end
            # 表示形式（h：時間、h以外：分）
            if (time_type == "h")
              val2 = schedule.getWorkTimeHours(work_time_unit)
            else
              val2 = schedule.getWorkTimeMinuts(work_time_unit)
            end
            base_cell.Offset(0, colOffset.to_i + 8).Value = val1 + val2
          end
        end
        
        if (idx < 0)
          addIdx += 1
        end
      else
        # エラーデータ。
        base_cell = sheet.Range(error_base_cell_str).Offset(addIdx + errIdx, 0)
        # エラー用セルに期間とタイトルを出力。
        base_cell.Offset(0, 0).Value = schedule.getTermString() 
        base_cell.Offset(0, 5).Value = schedule.getTitle()
        errIdx += 1
      end
    end
    
    if addIdx == 0
      return ;
    end
    # 合計列。
    # 横軸（日付）ごとの合計式を設定。
    work_time_base_address = @setting["BaseCell_WorkTime"]

    base_cell = sheet.Range(work_time_base_address).Offset(addIdx, 0)
    for curr_date in @startDate..@endDate
      offset = (curr_date - @startDate).to_i
      
      curr_cell = base_cell.Offset(0, offset)
      address_col1 = UtilExcel.getColumnAlphabet(curr_cell.Column())
      address_row1 = sheet.Range(work_time_base_address).Row()
      address = sprintf("%s%s", address_col1, address_row1)
      base_cell.Offset(0, offset).Formula = "=SUM(" + address + ":INDIRECT(ADDRESS(ROW()-1, COLUMN())))"
      
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
    # テンプレートからファイル作成。
    fso = WIN32OLE.new('Scripting.FileSystemObject')
    
    template_base_path = 'app/assets/excel/'
    template_abspath = template_base_path + '/template.xls'
    template_fullpath = fso.GetAbsolutePathName(template_abspath)
    return template_fullpath 
  end
  
  def getExcelWorkbook(scheduleList)
    # Excelのプロセスを取得する。
    @excelapp = getExcelObj()
    
    # テンプレートからファイル作成。
    template_fullpath = getTemplateFileFullpath()
    time_string = Time.now.strftime("%Y%m%d_%H%M%S")
    dest = UtilFile.getParentPath(template_fullpath) + '/output' + time_string + '.xls'
    FileUtils.copy(template_fullpath, dest)
    # ファイルを開く。
    book = @excelapp.workbooks.add(dest)
    
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
      autoFillColumnWidth(sheet)
    end
    
    # 全メンバーのスケジュールのみを取得。
    # シート作成。
    sheet = createMemberSheet(book, 'ALL')
    # 日付列出力。
    outputDateTerm(sheet)
    # シートにスケジュール出力。
    outputSchedule(sheet, scheduleList.getList())
    # 列幅調整。
    autoFillColumnWidth(sheet)
      
    @excelapp.displayAlerts = false
    # book.SaveAs('output' + time_string + '.xlsx') # フルパスでエラー発生？
    # 「56：Excel::XlExcel8」
    book.SaveAs({'Filename' => dest, 'FileFormat' => 56})
    @excelapp.displayAlerts = true
    
    # 渡す際は閉じなくてよいはず。
    # book.Close() TODO
    
    return book;
  end
  private :getExcelWorkbook
  
  def loadSetting(book)
    @setting = {}
    sheet = book.worksheets("設定")
    i = 4
    while true do
      curr = sheet.Range("C" + i.to_s)
      if (curr.value == nil || curr.value == "" ) 
        # 終了日を超えたらbreak。
        break
      end
      
      name = curr.value
      value = curr.Offset(0, 1).value
      @setting[name] = value
      i += 1
    end
  end
  private :loadSetting
  
  def autoFillColumnWidth(sheet)
    term = (@endDate - @startDate).to_i
    lastColumnAlpha = UtilExcel.getColumnAlphabet_onBase("K", term)
    sheet.Columns("K:" + lastColumnAlpha).AutoFit
  end
  
  def createMemberSheet(book, outputMember)
    sheet = book.Worksheets('templateSheet')
    # テンプレート用シートの前へコピー。
    sheet.Copy "before" => sheet
    # コピー先のシート名を変更
    @excelapp.ActiveSheet.Name = outputMember
    
    return book.Worksheets(outputMember)
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
