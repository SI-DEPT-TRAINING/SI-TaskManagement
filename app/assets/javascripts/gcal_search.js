
$(document).ready(function() {
	setEventHandler();
});

function setEventHandler() {

	$( "#dateFrom" ).datepicker({ autoSize: true });
	$( "#dateTo" ).datepicker({ autoSize: true });

	$('.open').click(function() {
		var _acount = $('#acount');
		var _passqword = $('#passqword');
		var _dateFrom = $('#dateFrom');
		var _dateTo = $('#dateTo');
		var data = {acount: _acount.val(), passqword: _passqword.val(), dateFrom: _dateFrom.val(), dateTo: _dateTo.val()};

	    $.ajax({
		    type: 'post',
		    url: 'http://localhost:3000/taskManagement/ajaxSetSession',
		    data: data,
		    dataType: 'json',
		    success: modalOpen
	    });

	});
}

function modalOpen() {
	// model Open
	$('#modal').dialog({
		modal: true,
		title: 'CSV UpLoad!!'
	});
}
