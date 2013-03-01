# encoding: utf-8
require './app/models/ScheduleModel.rb'

describe ScheduleModel do
    
    context "Googleカレンダを引数にNewした際" do
        before do
            # スタブを用意する。 
            @obj = Object.new()
            @obj.stub!(:title).and_return('たいとる')
            @obj.stub!(:desc).and_return("A001,B002,C003\r\ntest")
            @obj.stub!(:where).and_return('ばしょ')
            @obj.stub!(:st).and_return(Time.local(2013, 1, 20, 15, 00, 00))
            @obj.stub!(:en).and_return(Time.local(2013, 1, 20, 15, 45, 00))
        end
        
        context 'メソッド呼び出しで正しい値を返すこと。' do
            it 'getTitleメソッド呼び出しでタイトルを返すこと。' do
                model = ScheduleModel.new(@obj)
                model.getTitle().should eq 'たいとる'
            end
            it 'getDescメソッド呼び出しで説明を返すこと。' do
                model = ScheduleModel.new(@obj)
                model.getDesc().should eq "A001,B002,C003\r\ntest"
            end
            it 'getWhereメソッド呼び出しで場所を返すこと。' do
                model = ScheduleModel.new(@obj)
                model.getWhere().should eq 'ばしょ'
            end
            it 'getStartDateメソッド呼び出しで場所を返すこと。' do
                model = ScheduleModel.new(@obj)
                model.getStartDate().should eq Time.local(2013, 1, 20, 15, 00, 00)
            end
            it 'getEndDateメソッド呼び出しで場所を返すこと。' do
                model = ScheduleModel.new(@obj)
                model.getEndDate().should eq Time.local(2013, 1, 20, 15, 45, 00)
                
                @obj.stub!(:en).and_return(Time.local(2005, 8, 13, 22, 05, 30))
                model.getEndDate().should eq Time.local(2005, 8, 13, 22, 05, 30)
            end
        end
        context 'メソッド呼び出しで正しい値を返すこと。' do
            it 'getDesc_FirstLineメソッド呼び出しで説明の1行目を返すこと。' do
                model = ScheduleModel.new(@obj)
                model.getDesc_FirstLine().should eq "A001,B002,C003"
            end
            it 'getSectionListメソッド呼び出しで区分リストを返すこと。' do
                model = ScheduleModel.new(@obj)
                model.getSectionList().should eq ["A001", "B002", "C003"]
            end
        end
        context 'getWorkTimeHoursメソッドを呼び出した際' do
            it '引数なしで呼び出して工数（時間）を返すこと。' do
                model = ScheduleModel.new(@obj)
                model.getWorkTimeHours().should eq 0.75
            end
            it '引数15で呼び出して、15分単位の工数（時間）を返すこと。' do
                model = ScheduleModel.new(@obj)
                model.getWorkTimeHours(15).should eq 0.75
            end
            it '引数30で呼び出して、30分単位の工数（時間）を返すこと。' do
                model = ScheduleModel.new(@obj)
                model.getWorkTimeHours(30).should eq 0.5
            end
        end
        context 'getWorkTimeMinutsメソッドを呼び出した際' do
            it '引数なしで呼び出して工数（時間）を返すこと。' do
                @obj.stub!(:st).and_return(Time.local(2013, 1, 20, 15, 30, 00))
                @obj.stub!(:en).and_return(Time.local(2013, 1, 20, 16, 15, 00))
                model = ScheduleModel.new(@obj)
                model.getWorkTimeMinuts().should eq 45
            end
            it '引数15で呼び出して、15分単位の工数（分）を返すこと。' do
                @obj.stub!(:st).and_return(Time.local(2013, 1, 20, 15, 30, 00))
                @obj.stub!(:en).and_return(Time.local(2013, 1, 20, 16, 15, 00))
                model = ScheduleModel.new(@obj)
                model.getWorkTimeMinuts(15).should eq 45
            end
            it '引数20で呼び出して、20分単位の工数（分）を返すこと。' do
                @obj.stub!(:st).and_return(Time.local(2013, 1, 20, 15, 30, 00))
                @obj.stub!(:en).and_return(Time.local(2013, 1, 20, 16, 15, 00))
                model = ScheduleModel.new(@obj)
                model.getWorkTimeMinuts(20).should eq 40
            end
            it '引数30で呼び出して、30分単位の工数（分）を返すこと。' do
                @obj.stub!(:st).and_return(Time.local(2013, 1, 20, 15, 30, 00))
                @obj.stub!(:en).and_return(Time.local(2013, 1, 20, 16, 15, 00))
                model = ScheduleModel.new(@obj)
                model.getWorkTimeMinuts(30).should eq 30
            end
        end
        context 'getTermStringメソッドを呼び出した際' do
            it 'メソッド呼び出しで、期間の文字列を返すこと。' do
                model = ScheduleModel.new(@obj)
                model.getTermString().should eq "2013/01/20  15:00 - 15:45"
            end
        end
        context 'isProcessTarget?メソッドを呼び出した際' do
            context '説明の1行目が「A001,B002,C001」の際' do
                it '真を返す。' do
                    model = ScheduleModel.new(@obj)
                    model.isProcessTarget?().should be_true
                end
            end
            context '説明の1行目が「A001,B002test」の際' do
                it '偽を返す。' do
                    @obj.stub!(:desc).and_return("A001,B002test")
                    model = ScheduleModel.new(@obj)
                    model.isProcessTarget?().should be_false
                end
            end
        end
    end
end
