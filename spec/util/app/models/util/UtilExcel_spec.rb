# encoding: utf-8
require './app/models/util/UtilExcel.rb'

describe UtilExcel do
    
    describe 'getColumnNumberメソッド（Excelの列文字から列インデックスを返す）' do
        
        context 'メソッドを実行した時（小文字引数）' do
            
            it '引数に「a」を渡すと1を返す。' do
                UtilExcel.getColumnNumber("a").should equal 1
            end
            it '引数に「b」を渡すと2を返す。' do
                UtilExcel.getColumnNumber("b").should equal 2
            end
            it '引数に「c」を渡すと3を返す。' do
                UtilExcel.getColumnNumber("c").should equal 3
            end
            it '引数に「z」を渡すと26を返す。' do
                UtilExcel.getColumnNumber("z").should equal 26
            end
            it '引数に「aa」を渡すと27を返す。' do
                UtilExcel.getColumnNumber("aa").should equal 27
            end
            it '引数に「az」を渡すと52を返す。' do
                UtilExcel.getColumnNumber("az").should equal 52
            end
            it '引数に「ba」を渡すと53を返す。' do
                UtilExcel.getColumnNumber("ba").should equal 53
            end
        end
        
        context 'メソッドを実行した時（大文字引数）' do
            
            it '引数に「A」を渡すと1を返す。' do
                UtilExcel.getColumnNumber("A").should equal 1
            end
            it '引数に「B」を渡すと2を返す。' do
                UtilExcel.getColumnNumber("B").should equal 2
            end
            it '引数に「C」を渡すと3を返す。' do
                UtilExcel.getColumnNumber("C").should equal 3
            end
            it '引数に「Z」を渡すと26を返す。' do
                UtilExcel.getColumnNumber("Z").should equal 26
            end
            it '引数に「AA」を渡すと27を返す。' do
                UtilExcel.getColumnNumber("AA").should equal 27
            end
            it '引数に「AZ」を渡すと52を返す。' do
                UtilExcel.getColumnNumber("AZ").should equal 52
            end
            it '引数に「BA」を渡すと53を返す。' do
                UtilExcel.getColumnNumber("BA").should equal 53
            end
        end
        
        context 'メソッドを実行した時（大小文字混在の引数）' do
            
            it '引数に「Aa」を渡すと27を返す。' do
                UtilExcel.getColumnNumber("Aa").should equal 27
            end
            it '引数に「aA」を渡すと27を返す。' do
                UtilExcel.getColumnNumber("aA").should equal 27
            end
            it '引数に「Az」を渡すと52を返す。' do
                UtilExcel.getColumnNumber("Az").should equal 52
            end
            it '引数に「aZ」を渡すと52を返す。' do
                UtilExcel.getColumnNumber("aZ").should equal 52
            end
            it '引数に「Ba」を渡すと53を返す。' do
                UtilExcel.getColumnNumber("Ba").should equal 53
            end
            it '引数に「bA」を渡すと53を返す。' do
                UtilExcel.getColumnNumber("bA").should equal 53
            end
        end
        
        context 'メソッドを実行した時（Excelの列文字でない引数）' do
            
            it '引数に「A2」を渡すと-1を返す。' do
                UtilExcel.getColumnNumber("A2").should equal -1
            end
            it '引数に「aA」を渡すと-1を返す。' do
                UtilExcel.getColumnNumber("3").should equal -1
            end
            it '引数に「（空文字）」を渡すと-1を返す。' do
                UtilExcel.getColumnNumber("").should equal -1
            end
            it '引数に「nil」を渡すと-1を返す。' do
                UtilExcel.getColumnNumber(nil).should equal -1
            end
        end
    end
      
    describe 'getColumnAlphabet_onBaseメソッド' do
        context 'メソッドを実行した時（小文字引数）' do
            it '引数の列位置「Y」を起点とした位置を返す。' do
                UtilExcel.getColumnAlphabet_onBase("y", -1).should eq "X"
                UtilExcel.getColumnAlphabet_onBase("y", 0).should eq "Y"
                UtilExcel.getColumnAlphabet_onBase("y", 1).should eq "Z"
                UtilExcel.getColumnAlphabet_onBase("y", 2).should eq "AA"
                UtilExcel.getColumnAlphabet_onBase("y", 3).should eq "AB"
            end
        end
        
        context 'メソッドを実行した時（大文字引数）' do
            it '引数の列位置「AB」を起点とした位置を返す。' do
                UtilExcel.getColumnAlphabet_onBase("AB", -2).should eq "Z"
                UtilExcel.getColumnAlphabet_onBase("AB", -1).should eq "AA"
                UtilExcel.getColumnAlphabet_onBase("AB", 0).should eq "AB"
                UtilExcel.getColumnAlphabet_onBase("AB", 1).should eq "AC"
                UtilExcel.getColumnAlphabet_onBase("AB", 2).should eq "AD"
            end
        end
        
        context 'メソッドを実行した時（大小文字混在の引数）' do
            it '引数の列位置「CA」を起点とした位置を返す。' do
                UtilExcel.getColumnAlphabet_onBase("cA", -2).should eq "BY"
                UtilExcel.getColumnAlphabet_onBase("cA", -1).should eq "BZ"
                UtilExcel.getColumnAlphabet_onBase("cA", 0).should eq "CA"
                UtilExcel.getColumnAlphabet_onBase("cA", 1).should eq "CB"
                UtilExcel.getColumnAlphabet_onBase("cA", 2).should eq "CC"
            end
        end
        
        context 'メソッドを実行した時（Excelの列文字でない引数）' do
            it 'nilを返す。' do
                UtilExcel.getColumnAlphabet_onBase("a2", -2).should eq nil
                UtilExcel.getColumnAlphabet_onBase("3", -1).should eq nil
                UtilExcel.getColumnAlphabet_onBase("#", 0).should eq nil
                UtilExcel.getColumnAlphabet_onBase(" S", 1).should eq nil
                # UtilExcel.getColumnAlphabet_onBase("D ", 2).should eq nil
            end
        end
    end
      
    describe 'getColumnAlphabetメソッド' do
        context 'メソッドを実行した時（数値引数）' do
            it '引数に1を渡すと「A」を返す。' do
                UtilExcel.getColumnAlphabet(1).should eq "A"
            end
            it '引数に2を渡すと「B」を返す。' do
                UtilExcel.getColumnAlphabet(2).should eq "B"
            end
            it '引数に26を渡すと「Z」を返す。' do
                UtilExcel.getColumnAlphabet(26).should eq "Z"
            end
            it '引数に27を渡すと「AA」を返す。' do
                UtilExcel.getColumnAlphabet(27).should eq "AA"
            end
            it '引数に52を渡すと「AZ」を返す。' do
                UtilExcel.getColumnAlphabet(52).should eq "AZ"
            end
            it '引数に53を渡すと「BA」を返す。' do
                UtilExcel.getColumnAlphabet(53).should eq "BA"
            end
        end
        
        context 'メソッドを実行した時（ゼロ以下の引数）' do
            it '引数に0を渡すと「nil」を返す。' do
                UtilExcel.getColumnAlphabet(0).should eq nil
            end
            it '引数に負数を渡すと「nil」を返す。' do
                UtilExcel.getColumnAlphabet(-1).should eq nil
                UtilExcel.getColumnAlphabet(-2).should eq nil
                UtilExcel.getColumnAlphabet(-99).should eq nil
            end
        end
    end
      
    describe 'isCellAddress?メソッド' do
        context 'メソッドを実行した時（Excelのセルアドレスの引数）' do
            it '引数に「D7」を渡すとtrueを返す。' do
                UtilExcel.isCellAddress?('D7').should be_true
            end
            it '引数に「D7：H15」を渡すと配列trueを返す。' do
                UtilExcel.isCellAddress?('D7:H15').should be_true
            end
        end
        
        context 'メソッドを実行した時（Excelのセルアドレスでない引数）' do
            it '引数に「D7：」を渡すとを返す。' do
                UtilExcel.isCellAddress?('D7:').should be_false
            end
            it '引数に「:F4」を渡すとfalseを返す。' do
                UtilExcel.isCellAddress?(':F4').should be_false
            end
            it '引数に「A3:F4:H6」を渡すとfalseを返す。' do
                UtilExcel.isCellAddress?('A3:F4:H6').should be_false
            end
            it '引数にnilを渡すとfalseを返す。' do
                UtilExcel.isCellAddress?(nil).should be_false
            end
        end
    end
    
    describe 'getCornerメソッド' do
        context 'メソッドを実行した時（Excelのセルアドレスの引数）' do
            it '引数に「D7」を渡すと配列["D", "7", "D", "7"]を返す。' do
                UtilExcel.getCorner('D7').should eq ["D", "7", "D", "7"]
            end
            it '引数に「D7：H15」を渡すと配列["D", "7", "H", "15"]を返す。' do
                UtilExcel.getCorner('D7:H15').should eq ["D", "7", "H", "15"]
            end
        end
        
        context 'メソッドを実行した時（Excelのセルアドレスでない引数）' do
            it '引数に「D7：」を渡すとnilを返す。' do
                UtilExcel.getCorner('D7:').should eq nil
            end
            it '引数に「:F4」を渡すとnilを返す。' do
                UtilExcel.getCorner(':F4').should eq nil
            end
            it '引数に「A3:F4:H6」を渡すとnilを返す。' do
                UtilExcel.getCorner('A3:F4:H6').should eq nil
            end
            it '引数にnilを渡すとnilを返す。' do
                UtilExcel.getCorner(nil).should eq nil
            end
        end
    end
    
    describe 'getRowsAddressメソッド' do
        context 'メソッドを実行した時（Excelのセルアドレスの引数）' do
            it '引数に「D7」を渡すと"7:7"を返す。' do
                UtilExcel.getRowsAddress('D7').should eq "7:7"
            end
            it '引数に「D7：H15」を渡すと"7:15"を返す。' do
                UtilExcel.getRowsAddress('D7:H15').should eq "7:15"
            end
        end
        
        context 'メソッドを実行した時（Excelのセルアドレスでない引数）' do
            it '引数に「D7：」を渡すとnilを返す。' do
                UtilExcel.getRowsAddress('D7:').should eq nil
            end
            it '引数に「:F4」を渡すとnilを返す。' do
                UtilExcel.getRowsAddress(':F4').should eq nil
            end
            it '引数に「A3:F4:H6」を渡すとnilを返す。' do
                UtilExcel.getRowsAddress('A3:F4:H6').should eq nil
            end
            it '引数にnilを渡すとnilを返す。' do
                UtilExcel.getRowsAddress(nil).should eq nil
            end
        end
    end
    
    describe 'getColumnsAddressメソッド' do
        context 'メソッドを実行した時（Excelのセルアドレスの引数）' do
            it '引数に「D7」を渡すと"7:7"を返す。' do
                UtilExcel.getColumnsAddress('D7').should eq "D:D"
            end
            it '引数に「D7：H15」を渡すと"7:15"を返す。' do
                UtilExcel.getColumnsAddress('D7:H15').should eq "D:H"
            end
        end
        
        context 'メソッドを実行した時（Excelのセルアドレスでない引数）' do
            it '引数に「D7：」を渡すとnilを返す。' do
                UtilExcel.getColumnsAddress('D7:').should eq nil
            end
            it '引数に「:F4」を渡すとnilを返す。' do
                UtilExcel.getColumnsAddress(':F4').should eq nil
            end
            it '引数に「A3:F4:H6」を渡すとnilを返す。' do
                UtilExcel.getColumnsAddress('A3:F4:H6').should eq nil
            end
            it '引数にnilを渡すとnilを返す。' do
                UtilExcel.getColumnsAddress(nil).should eq nil
            end
        end
    end
    
    describe 'getTopLeftAddressメソッド' do
        context 'メソッドを実行した時（Excelのセルアドレスの引数）' do
            it '引数に「D7」を渡すと"D7"を返す。' do
                UtilExcel.getTopLeftAddress('D7').should eq "D7"
            end
            it '引数に「D7：H15」を渡すと"D7"を返す。' do
                UtilExcel.getTopLeftAddress('D7:H15').should eq "D7"
            end
        end
        
        context 'メソッドを実行した時（Excelのセルアドレスでない引数）' do
            it '引数に「D7：」を渡すとnilを返す。' do
                UtilExcel.getTopLeftAddress('D7:').should eq nil
            end
            it '引数に「:F4」を渡すとnilを返す。' do
                UtilExcel.getTopLeftAddress(':F4').should eq nil
            end
            it '引数に「A3:F4:H6」を渡すとnilを返す。' do
                UtilExcel.getTopLeftAddress('A3:F4:H6').should eq nil
            end
            it '引数にnilを渡すとnilを返す。' do
                UtilExcel.getTopLeftAddress(nil).should eq nil
            end
        end
    end
    
    describe 'getBottomRightAddressメソッド' do
        context 'メソッドを実行した時（Excelのセルアドレスの引数）' do
            it '引数に「D7」を渡すと"D7"を返す。' do
                UtilExcel.getBottomRightAddress('D7').should eq "D7"
            end
            it '引数に「D7：H15」を渡すと"H15"を返す。' do
                UtilExcel.getBottomRightAddress('D7:H15').should eq "H15"
            end
        end
        
        context 'メソッドを実行した時（Excelのセルアドレスでない引数）' do
            it '引数に「D7：」を渡すとnilを返す。' do
                UtilExcel.getBottomRightAddress('D7:').should eq nil
            end
            it '引数に「:F4」を渡すとnilを返す。' do
                UtilExcel.getBottomRightAddress(':F4').should eq nil
            end
            it '引数に「A3:F4:H6」を渡すとnilを返す。' do
                UtilExcel.getBottomRightAddress('A3:F4:H6').should eq nil
            end
            it '引数にnilを渡すとnilを返す。' do
                UtilExcel.getBottomRightAddress(nil).should eq nil
            end
        end
    end
end
