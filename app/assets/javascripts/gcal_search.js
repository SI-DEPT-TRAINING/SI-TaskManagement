
//初期表示
$(document).ready(function() {
	setEventHandler();
});

// イベントハンドラー
function setEventHandler() {

	// datepickerを設定
	$( "#dateFrom" ).datepicker({ autoSize: true });
	$( "#dateTo" ).datepicker({ autoSize: true });

	// Ajaxをクリックイベント設定
	$('.open').click(function() {
		var _acount = $('#acount');
		var _password = $('#password');
		var _dateFrom = $('#dateFrom');
		var _dateTo = $('#dateTo');
		var csrf_token = $("meta[name=csrf-token]").attr("content");
		var data = {acount: _acount.val(), password: _password.val(), dateFrom: _dateFrom.val(), dateTo: _dateTo.val()};

		// URLは環境依存、TODO
		// jQuery_Ajax_通信実行
	    $.ajax({
		    type: 'post',
		    url: '/taskManagement/ajaxSetSession',
		    data: data,
		    dataType: 'json',
		    success: modalOpen
	    });
	});
}

// モーダルオープン
function modalOpen(data) {

	// 通信結果のエラーハンドリングを追加する　TODO
	$('#modal').dialog({
		modal: true,
		title: 'CSV UpLoad!!'
	});
}
