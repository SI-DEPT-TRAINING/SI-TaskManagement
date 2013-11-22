# -*- encoding: utf-8 -*-
class UtilFile

  def self.getFilenameFullPath(file_abspath)
    # ファイル名（フルパス）の取得。
    template_fullpath = fso.GetAbsolutePathName(file_abspath)
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

end
