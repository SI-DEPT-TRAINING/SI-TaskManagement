# -*- coding: utf-8 -*-
require 'win32ole'

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
    # データのマージ。
    # データ書き込み。
    getExcelWorkbook(scheduleList);
    
    
    return '';
  end
  
  def getExcelScheduleList()
    #デバック用（本来は画面表示用のオブジェクトへ詰め替える予定）
    cnt = 1;
    scheduleList = ScheduleList.new()
    @googleEventList.each do |serchOutputModel|
      serchOutputModel.eventList.each do |event|
        puts Kconv.tosjis(serchOutputModel.name)
        puts Kconv.tosjis(serchOutputModel.acount)
        #タイトル
        #puts event.title
        #詳細
        #puts event.desc
        #場所
        #puts event.where
        #開始日時
        #puts event.st
        #終了日時
        #puts event.en
        
        schedule = ScheduleModel.new(event);
        schedule.username = serchOutputModel.name
        scheduleList.add(schedule)
        # ↑ここまで取得処理。
            
        #puts(Kconv.tosjis(schedule.getTitle()))
        #puts schedule.getTitle().encode("Shift_JIS")
        puts schedule.getDesc()
        puts schedule.getWhere()
        puts schedule.getStartDate()
        puts schedule.getEndDate()
        
        print(cnt , "件目".encode("Shift_JIS"))
        cnt = cnt + 1
      end
    end
    return scheduleList;
  end
  private :getExcelScheduleList
  
  def outputDateTerm(sheet)
    curr = @startDate
    base_cell = sheet.Range("K12")
    del_row_number = base_cell.Row().to_s
    de = base_cell.Column().to_s
    #sheet.Range(del_row_number + ":" + del_row_number).Delete
    p @setting["InitDataCellRange"]
#    /(\$([A-Z]+)\$([0-9]+))(:(\$([A-Z]+)\$([0-9]+)))?/ =~ base_cell.Address()
    /(\$?([A-Z]+)\$?([0-9]+))(:(\$?([A-Z]+)\$?([0-9]+)))?/ =~ @setting["InitDataCellRange"]
    address_col1 = $2
    address_row1 = $3
    address_col2 = $6
    address_row2 = $7
    p address_col1
    p address_row1
    p address_col2
    p address_row2
    
    delele_col = sprintf("%s:%s", $2, $6)
    delele_row = sprintf("%s:%s", $3, $7)
    p delele_col
    p delele_row
#    sheet.Range(address_col + ":" + address_col).Delete
 #   sheet.Range(address_row + ":" + address_row).Delete
    sheet.Range(delele_col).Delete
    sheet.Range(delele_row).Delete
  #  sheet.Range(address_row + ":" + address_row).Delete
    
    error_base_cell_str = @setting["ErrorBaseCell"]
    error_base_cell = sheet.Range(error_base_cell_str)
    while true do
      base_cell = sheet.Range("K12")
      if (@endDate - curr) < 0
        # 終了日を超えたらbreak。
        break
      end
      
      offset = (curr - @startDate)
      offset = offset.to_i
      puts offset
      
      ins_ = base_cell.Offset(0, offset).Row().to_s
      /(\$([A-Z]+)\$([0-9]+))(:(\$([A-Z]+)\$([0-9]+)))?/ =~ base_cell.Offset(0, offset).Address()
      tmp_col = $2
      tmp_row = $3
      puts tmp_col + ":" + tmp_col
      sheet.Range(tmp_col + ":" + tmp_col).Insert()
      puts tmp_col + ":" + tmp_col
      puts base_cell.Address()
      puts offset
      base_cell.Offset(0, offset).Address()
      puts
      
      base_cell = sheet.Range("K12")
      # 日付列の入力。
      base_cell.Offset(-4, offset).Value = curr.to_s
      base_cell.Offset(-4, offset).NumberFormatLocal =  'm"月"d"日"'
      # 曜日列の入力。
      base_cell.Offset(-3, offset).Formula = '=' + base_cell.Offset(-4, offset).Address() 
      base_cell.Offset(-3, offset).NumberFormatLocal =  'aaa' #　「"aaa"：曜日」
      # 列幅調整。
      sheet.Columns(10+offset).AutoFit
      sheet.Rows(8).AutoFit
      
      # 次の日へ
      curr += 1
    end
  end
  private :outputDateTerm
  
  def outputSchedule(sheet, scheduleList)
    addIdx = 0
    errIdx = 0
    #allIdx = 0
    # base_cell = sheet.Range("C12")
    error_base_cell_str = @setting["ErrorBaseCell"]
    work_time_unit = @setting["WorkTimeUnit"].to_i
    time_type = @setting["TimeType"]
    scheduleList.each do |schedule|
      # スケジュールデータか確認する。
      flg = schedule.isProcessTarget?()
      if (flg == true) 
        # スケジュールデータ。
        base_cell = sheet.Range("C12").Offset(addIdx, 0)
        # 分類のリストを取得。
        sectionList = schedule.getSectionList()
        #sheet.Range("C12:C").Offset(addIdx, 0)
        #schedule.isSameSection?(scheduleList)
        idx = -1
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
          base_cell = sheet.Range("C12").Offset(idx, 0)
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
        #base_cell.Offset(0, -1).Value = flg
        #base_cell.Offset(addIdx, 4).Value = schedule.getTitle()
        #base_cell.Offset(addIdx, 5).Value = schedule.getWhere()
        #base_cell.Offset(addIdx, -0).Value = schedule.getStartDate()
        #base_cell.Offset(0, 7).Value = schedule.getWorkTimeMinuts()
        #base_cell.Offset(addIdx, 7).Value = schedule.getEndDate()
        
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
#      allIdx += 1
    end
    
    if addIdx == 0
      return ;
    end
    # 合計列。
    # 横軸（日付）ごとの合計式を設定。
    /(\$?([A-Z]+)\$?([0-9]+))(:(\$?([A-Z]+)\$?([0-9]+)))?/ =~ @setting["InitDataCellRange"]
    address_cell1 = $1

    base_cell = sheet.Range(address_cell1).Offset(addIdx, 0)
    curr = @startDate
    while true do
      if (@endDate + 1 - curr) < 0
        # 終了日を超えたらbreak。
        break
      end
      
      offset = (curr - @startDate)
      offset = offset.to_i
      
      # base_cell.Offset(0, offset).Value = "123"
      /(\$?([A-Z]+)\$?([0-9]+))(:(\$?([A-Z]+)\$?([0-9]+)))?/ =~ base_cell.Offset(0, offset).Address()
      address_col1 = $2
      base_cell.Offset(0, offset).Formula = "=SUM(" + address_col1 + "12:INDIRECT(ADDRESS(ROW()-1, COLUMN())))"
      p base_cell.Offset(0, offset).Address()
      
      # 次の日へ
      curr += 1
    end
    
    #base_cell.Offset(0, 0).Formula = "=SUM(K12:INDIRECT(ADDRESS(ROW()-1, COLUMN())))"
  end
  
  def getExcelObj()
    begin
      # 既存のExcelプロセスが起動していれば、それを使いまわす。
      excelapp = WIN32OLE::connect("Excel.Application")
    rescue WIN32OLERuntimeError 
      # Excelプロセスがないので、新たにプロセスを作成。
      excelapp = WIN32OLE.new("Excel.Application")
    end
    return excelapp 
  end
  
  def getTemplateFileFullpath()
    # テンプレートからファイル作成。
    fso = WIN32OLE.new('Scripting.FileSystemObject')
    
    template_base_path = 'app/controllers/tmp/'
    template_abspath = template_base_path + '/template.xls'
    template_fullpath = fso.GetAbsolutePathName(template_abspath)
    return template_fullpath 
  end
  
  def getExcelWorkbook(scheduleList)
    # Excelのプロセスを取得する。
    @excelapp = getExcelObj()
    
    # テンプレートからファイル作成。
    template_fullpath = getTemplateFileFullpath()
    book = @excelapp.workbooks.add(template_fullpath)
    
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
    end
    
    fso = WIN32OLE.new('Scripting.FileSystemObject')
    time_string = Time.now.strftime("%Y%m%d_%H%M%S")
    template_base_path = 'app/controllers/tmp/'
    output_abspath = template_base_path + '/output' + time_string + '.xlsx'
    output_fullpath = fso.GetAbsolutePathName(output_abspath)
    # book.SaveAs('output' + time_string + '.xlsx') # フルパスでエラー発生？
    book.SaveAs(output_fullpath) # フルパスでエラー発生？
    filname = 'C:/ProgramFiles/eclipse/ruby_work/SI-TaskManagement/app/output.xlsx'
    #book.SaveAs(filname)
    #book.Save(fso.GetAbsolutePathName(xlsfilename))
    
    @excelapp.Workbooks.Close
    # 必ず終了する。でないとプロセスが残る。
    @excelapp.Quit
    
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
  def createMemberSheet(book, outputMember)
    sheet = book.Worksheets('templateSheet')
    # テンプレート用シートの前へコピー。
    sheet.Copy "before" => sheet
    # コピー先のシート名を変更
    @excelapp.ActiveSheet.Name = outputMember
    puts outputMember.encode("Shift_JIS")
    
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
