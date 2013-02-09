# -*- coding: utf-8 -*-
require 'win32ole'

class CreateWorkbook

  # コンストラクタ。
  def initialize(googleEventList, fileName)
    @googleEventList = googleEventList
    @fileName = fileName
    @memberList = ["Aさん", "Bさん"] # テスト用
    @memberList = ["Aさん", "Bさん", "Cさん"] # テスト用
    @startDate = Date::new(2013, 1, 20) # テスト用
    @endDate = Date::new(2013, 1, 30) # テスト用
  end
  
  def doExe
    # データ取得。
    scheduleList = getExcelScheduleList()
    # データのマージ。
    abc = '123';
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
  
  def getExcelWorkbook(scheduleList)
    template_base_path = 'app/controllers/tmp/'
    # テンプレートからファイル作成。
    fso = WIN32OLE.new('Scripting.FileSystemObject')
    @excelapp = WIN32OLE.new('Excel.Application')
    template_abspath = template_base_path + '/template.xls'
    template_fullpath = fso.GetAbsolutePathName(template_abspath)
    book = @excelapp.workbooks.add(template_fullpath)
    @memberList.each do |outputMember|
      # シート作成。
      sheet = createMemberSheet(book, outputMember)
      curr = @startDate
      base_cell = sheet.Range("K12")
      del_row_number = base_cell.Row().to_s
      de = base_cell.Column().to_s
      #sheet.Range(del_row_number + ":" + del_row_number).Delete
      /(\$([A-Z]+)\$([0-9]+))(:(\$([A-Z]+)\$([0-9]+)))?/ =~ base_cell.Address()
      address_col = $2
      address_row = $3
#      address_col = base_cell.Address()[/[a-zA-Z]+/]
 #     address_row = base_cell.Address()[/[0-9]+/]
      sheet.Range(address_col + ":" + address_col).Delete
      sheet.Range(address_row + ":" + address_row).Delete
      
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
      
      memberScheduleList = scheduleList.narrowMember(outputMember)
      i = 0
      base_cell = sheet.Range("C12")
      memberScheduleList.each do |schedule|
        ins_ = base_cell.Offset(i, 0).Row().to_s
        sheet.Range(ins_ + ":" + ins_).Insert()
        base_cell = sheet.Range("C12")
        / *([0-9]+) *, *([0-9]+) *,  *([0-9]+) */ =~ schedule.getDesc()
        /([^,]*),([^,]*),([^,]*)/ =~ schedule.getDesc()
        sectionA = $1
        sectionB = $2
        sectionC = $3
        
        base_cell.Offset(i, 0).Value = "13" # とりあえず固定
        base_cell.Offset(i, 1).Value = sectionA
        base_cell.Offset(i, 2).Value = sectionB
        base_cell.Offset(i, 3).Value = sectionC
        base_cell.Offset(i, 4).Value = "適当（分類で決まる）" # とりあえず固定
        base_cell.Offset(i, 5).Value = "適当（分類で決まる）" # とりあえず固定
        base_cell.Offset(i, 6).Value = "適当（分類で決まる）" # とりあえず固定
        
        scheDateStr = schedule.getStartDate()
        scheDate = Date::new(scheDateStr.year, scheDateStr.month, scheDateStr.day)
        colOffset = scheDate - @startDate
        if (0 <= colOffset) #後後いらない。
          if (colOffset <= (@endDate - @startDate)) #後後いらない。
            base_cell.Offset(i, colOffset.to_i + 8).Value = schedule.getWorkTimeMinuts()
          end
        end
        p colOffset
        #base_cell.Offset(i, 3).Value = schedule.isProcessTarget()
        #base_cell.Offset(i, 4).Value = schedule.getTitle()
        #base_cell.Offset(i, 5).Value = schedule.getWhere()
        base_cell.Offset(i, 6).Value = schedule.getStartDate()
        base_cell.Offset(i, 7).Value = schedule.getWorkTimeMinuts()
        #base_cell.Offset(i, 7).Value = schedule.getEndDate()
        i=i+1
      end
    end
    
    time_string = Time.now.strftime("%Y%m%d_%H%M%S")
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
end
