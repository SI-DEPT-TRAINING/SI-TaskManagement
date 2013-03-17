# -*- encoding: utf-8 -*-
=begin 
バリデートモジュール群を定義する
=end
module ValidetionModule

=begin 
必須チェック
=end
  class MustChecker
    
      @error = nil
      @errorMsg = nil
      # ----------------------------
      # 必須チェック
      # ----------------------------
      def initialize(id, value)
        msgStr = "の値は必須です。"
        if value.blank?
          @error  = true
          @errorMsg = id + msgStr
          return
        end
      end
      attr_accessor :error
      attr_accessor :errorMsg
  end

=begin 
日付関連のチェッカーです
=end
require 'date'
  class DateChecker < MustChecker

      def initialize(id, value)
        super(id, value)
        
        unless @error then
          format(id, value)
        end
        
        unless @error then
          existence(id, value)
        end

      end

      # ----------------------------
      # フォマットチェック　yyyy/mm/dd
      # ----------------------------
      def format(id, value)
        msgStr = "のフォーマットエラーです。yyyy/mm/ddの日付を入力してください。"
        temp = value.split("/")
        if temp.size != 3 then
          @error = true
          @errorMsg = id + msgStr
        end 
      end

      # ----------------------------
      # 日付の正当性チェック
      # ----------------------------
      def existence(id, value)
        msgStr = "には、正しい日付を設定してください。"
        temp = value.gsub("/", "")
        if temp.length != 8 then
          @error = true
          @errorMsg = id + msgStr
        end

        tempValid = Date.valid_date?(temp[0, 4].to_i, temp[4,2].to_i, temp[6,2].to_i)
        if tempValid.blank? then
          @error = true
          @errorMsg = id + msgStr
        end
      end

      # ----------------------------
      # 日付の正当性チェック
      # ----------------------------
      def diff(valueFrom, valueTo)
        msgStr = "日付Fromが日付Toより未来日付に設定されています。"
        tempFrom = valueFrom.gsub("/", "")
        tempTo  = valueTo.gsub("/", "")
        dFrom = Date.new(tempFrom[0, 4].to_i, tempFrom[4,2].to_i, tempFrom[6,2].to_i)
        dTo = Date.new(tempTo[0, 4].to_i, tempTo[4,2].to_i, tempTo[6,2].to_i)
        
        if (dTo - dFrom).to_i < 0 then
          @error = true
          @errorMsg = msgStr
        end
      end
      
      attr_accessor :error
      attr_accessor :errorMsg
  end

=begin 
CSVファイルチェッカー
=end
  class CsvChecker
    
      @error = nil
      @errorMsg = nil
      @line = nil

      # ----------------------------
      # 必須チェック
      # ----------------------------
      def initialize(csvFile)
        msgStr="CSVフォーマットエラーです。CSVファイルの中身を確認してください。"
        if csvFile.blank?
          @error  = true
          @errorMsg = msgStr
          return
        end
      end

      # ----------------------------
      # カラム数チェック
      # ----------------------------
      def rows(csvFile)
        msgStr="CSVフォーマットエラーです。CSVファイルの中身を確認してください。"
        # 配列 >>　レコード
        @line = csvFile.read.split("\n")
        # カラム >> 配列
        @line.each do |rows|
          row = rows.split(",")
          if row.size != 2 then
            @error = true
            @errorMsg = msgStr
            break
          end
        end
      end

      # ----------------------------
      # 拡張子チェック
      # ----------------------------
      def contentTyoe(csvFile)

        msgStr="アップロード可能なファイルはCSVファイルのみです。"
        name = csvFile.original_filename

        if !['.csv'].include?(File.extname(name).downcase)
          @error = true
          @errorMsg = msgStr
        end

      end

      attr_accessor :error
      attr_accessor :errorMsg
      attr_accessor :line
  end


end
