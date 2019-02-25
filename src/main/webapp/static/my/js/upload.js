$(function(e){
	var $list = $("#thelist");
	var fileSize = 0;  //总文件大小
	var fileName = []; //文件名列表
	var fileSizeOneByOne =[];//每个文件大小
	var uploader;// 实例化 
	uploader = WebUploader.create({
		auto:false, //是否自动上传
		swf: CTX + '/static/webuploader/Uploader.swf',
		pick: {
            id:"#multi",
            name:"multiFile",  //这个地方 name没什么用，虽然打开调试器，input的名字确实改过来了。但是提交到后台取不到文件。如果想自定义file的name属性，还是要和fileVal 配合使用。
            label: '点击选择图片'
        },
        fileVal:'multiFile',//和name属性配合使用
        server: CTX + '/autocode/upload',
        duplicate: true,//是否可重复选择同一文件
        resize: false,//图片压缩
        compress: null,//图片不压缩
        chunked: true,  //分片处理
        chunkSize: 5 * 1024 * 1024, //每片5M
        chunkRetry:false,//如果失败，则不重试
        threads:1,//上传并发数。允许同时最大上传进程数。
        formData: {
        	//"uid": 123,
            "status":"multi",
            "existFlg":'false'
        },
        accept: {//只能上传这些格式的文件
            title: 'Images',
            extensions: 'gif,jpg,jpeg,bmp,png'
        },
        //fileNumLimit:50,//验证文件总数量, 超出则不允许加入队列
        disableGlobalDnd: true // 禁掉全局的拖拽功能。这样不会出现图片拖进页面的时候，把图片打开。  
	});
	
	//当有文件添加进来的时候
	uploader.on("fileQueued", function(file) {
	    console.log("fileQueued:");
	    $list.append("<div id='"+ file.id + "' class='item'>" +
				        "<h4 class='info'>" + file.name + "</h4>" +
				        "<p class='state'>等待上传...</p>" +
				     "</div>");
	});
	
	//当开始上传流程时触发
    uploader.on("startUpload", function() {
      console.log("startUpload");
      //添加额外的表单参数
      $.extend(true, uploader.options.formData, {"fileSize":fileSize,"multiFileName":fileName,"fileSizeOneByOne":fileSizeOneByOne}); 
    });
	
	//当所有文件上传结束时触发
	uploader.on("uploadFinished",function(){
	    console.log("uploadFinished:");
	});

	//当文件上传成功时触发。
	uploader.on("uploadSuccess", function(file) {
	  $("#"+file.id).find("p.state").text("已上传");
	});
	//当文件上传失败时触发
	uploader.on("uploadError", function(file,reason) {
		console.log("uploadError:");
        $("#"+file.id).find("p.state").text("上传出错");
        console.log(file);
        console.log(reason);
        
        var fileArray = uploader.getFiles(); //多个文件
        for(var i = 0;i<fileArray.length;i++){
        	//取消文件上传
            uploader.cancelFile(fileArray[i]);
            //从队列中移除掉
            uploader.removeFile(fileArray[i],true);
        }
        //发生错误重置webupload,初始化变量
        uploader.reset();
        fileSize = 0;
        fileName = [];
        fileSizeOneByOne=[];
    });
	
	//当validate不通过时，会以派送错误事件的形式通知调用者
    uploader.on("error",function(){
	    uploader.reset();
	    fileSize = 0;
	    fileName = [];
	    fileSizeOneByOne=[];
	    alert("请上传图片格式的文件！");
    });
	
	//上传按钮绑定事件
	$("#multiUpload").on("click", function() {
         uploader.upload();
    });
	
	
	//选择文件之后执行上传  
    $(document).on("change","input[name='multiFile']", function() {
          //multiFileName
          var fileArray1 = uploader.getFiles();
          fileSize = 0;
          fileName = [];
          fileSizeOneByOne = [];
          for(var i = 0 ;i<fileArray1.length;i++){
              //fileNames.push(fileArray1[i].name); //input 框用
              //后台用
              fileSize += fileArray1[i].size;
              fileSizeOneByOne.push(fileArray1[i].size);
              fileName.push(fileArray1[i].name);
          }
          console.log(fileSize);
          console.log(fileSizeOneByOne);
          console.log(fileName);
    });
});

