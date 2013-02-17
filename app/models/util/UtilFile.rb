# -*- encoding: utf-8 -*-
# require "gcalapi";

class UtilFile
  # 
  @@fso = WIN32OLE.new('Scripting.FileSystemObject')
  
  def self.getFilenameFullPath(file_abspath)
    # ファイル名（フルパス）の取得。
    template_fullpath = @@fso.GetAbsolutePathName(file_abspath)
    return template_fullpath 
  end
  
  def self.getParentPath(file_path)
    # 親ディレクトリ名の取得。
    parentPath = File::dirname(file_path)
    return parentPath 
  end
  
  def self.copy(file1, file2)
    FileUtils.copy(file1, file2)
    return 
  end
  
  def self.getFilenameOnly(file_path)
    filename_only = File::basename(file_path)
    return filename_only 
  end
  
  def self.getExt(file_path)
    ext = File::extname(file_path)
    return ext 
  end
  
  def self.getColumnAlphabet(colNumber)
    if colNumber <= 0
      return nil
    end
    
    alpha_str = ""
    tmp = colNumber - 1
    while (true)
      ascii_int = "A"[0].to_i + (tmp % 26).to_i
      alpha = ascii_int.chr
      alpha_str = alpha.to_s + alpha_str
      tmp = tmp / 26
      tmp = tmp.to_i
      if tmp == 0
        break
      end
    end
    return alpha_str 
  end
  
end
