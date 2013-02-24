# encoding: utf-8
require './app/models/util/UtilFile.rb'

describe UtilFile do
    
    describe 'getFilenameFullPathメソッド' do
        context 'メソッドを実行した時（相対パスの引数）' do
            it '引数にファイル名を渡すとフルパスを返す。' do
                UtilFile.getFilenameFullPath('test.txt').should eq "C:\\ProgramFiles\\eclipse\\ruby_work\\SI-TaskManagement\\test.txt"
            end
            it '引数に相対ファイル名を渡すと絶対パスを返す。' do
                UtilFile.getFilenameFullPath('test\\test.txt').should eq "C:\\ProgramFiles\\eclipse\\ruby_work\\SI-TaskManagement\\test\\test.txt"
            end
        end
    end
    
    describe 'getParentPathメソッド' do
        context 'メソッドを実行した時（ファイルパスの引数）' do
            it '引数にファイル名を渡すと親のパスを返す。' do
                UtilFile.getParentPath('C:\\test1\\test2\\test.txt').should eq "C:\\test1\\test2"
            end
            it '引数にフォルダ名を渡すと親のパスを返す。' do
                UtilFile.getParentPath('C:\\test1\\test2').should eq "C:\\test1"
                UtilFile.getParentPath('C:\\test1\\test2\\').should eq "C:\\test1"
            end
            it '引数にルートを渡すとルートのパスを返す。' do
                UtilFile.getParentPath('C:\\test1').should eq "C:\\"
                UtilFile.getParentPath('C:\\').should eq "C:\\"
            end
        end
    end
    
    describe 'getFilenameOnlyメソッド' do
        context 'メソッドを実行した時（ファイルパスの引数）' do
            it '引数にフルパスファイル名を渡すとファイル名のみを返す。' do
                UtilFile.getFilenameOnly('C:\\test1\\test2\\test.txt').should eq "test.txt"
            end
            it '引数にフォルダパスを渡すとフォルダ名を返す。' do
                UtilFile.getFilenameOnly('C:\\test1\\test2').should eq "test2"
            end
        end
    end
    
    describe 'getExtメソッド' do
        context 'メソッドを実行した時（ファイルパスの引数）' do
            it '引数にフルパスファイル名を渡すとファイル名のみを返す。' do
                UtilFile.getExt('C:\\test1\\test2\\test.txt').should eq ".txt"
            end
            it '引数にフォルダパスを渡すとフォルダ名を返す。' do
                UtilFile.getExt('C:\\test1\\test2\\test').should eq ""
            end
            it '引数にフォルダパスを渡すとフォルダ名を返す。' do
                UtilFile.getExt('C:\\test1\\test2').should eq ""
            end
        end
    end
end
