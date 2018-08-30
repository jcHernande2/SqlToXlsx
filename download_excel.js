$(document).on('click', '.btn-download-excel', function(e){
	var btn_download=$(this);
	var id_form=$(this).parents("form").attr("id");
	var data_url=$(this).parents("form").attr("data-url");
	var view_result=$(this).parents("form").attr("data-view");
	var view_msg=$(this).parents("form").attr("data-msg");
	var formData=new FormData(document.getElementById(id_form));
	//inicio Progress Bar;
	$(btn_download).parents("form").find(".btn-search").attr('disabled','disabled');
	$(btn_download).attr("disabled","disabled");
	$(btn_download).html('<span class="glyphicon glyphicon-refresh"></span> Exportando..');
	$.ajax({
		url: data_url,
		type: 'POST',
		data: formData,
		processData: false,
		contentType: false,
		dataType: 'json',
		success: function(result){
			if(result.code){
				$('#'+view_msg).html(result.msg);
			}else{
				var $a = $("<a>");
				$a.attr("href",result.filedata);
				$("body").append($a);
				$a.attr("download",result.filename+".xlsx");
				$a[0].click();
				$a.remove();
				$(btn_download).removeAttr("disabled");
				$(btn_download).html('<span class="glyphicon glyphicon-file"></span> Exportar');
				$(btn_download).parents("form").find(".btn-search").removeAttr("disabled");
				//fin Progress Bar;
			}
		},
		error: function(XMLHttpRequest, textStatus, errorThrown) { 
			$('#'+view_msg).html('<div class="alert alert-danger"><button type="button" class="close" data-dismiss="alert">×</button>'+XMLHttpRequest.responseText+'</div>');
			$(btn_download).parents("form").find(".btn-search").removeAttr("disabled");
			//Fin Progress Bar;
		}
	});
	e.preventDefault();
});
