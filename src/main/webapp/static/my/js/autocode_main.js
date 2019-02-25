$(function(){
	$("#btn-commit").on("click",function(e){
		start();
	});
	$("#btn-reset").on("click",function(e){
		$("#inputForm").resetForm();
	});
});

function start(){
	$("#inputForm").attr("action",CTX+"/autocode/download");
	$("#inputForm").submit();
}