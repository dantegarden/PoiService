<%@ page language="java" contentType="text/html; charset=UTF-8"
	pageEncoding="UTF-8"%>
<%@include file="/WEB-INF/commons/taglibs.jsp"%>
<%@include file="/WEB-INF/commons/pre-include.jsp"%>

<!DOCTYPE html>
<html>
<head>
<base href="<%=basePath%>">

<title>DVT AUTOCODE WEB-UI</title>

<meta charset="UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1.0" />
<link rel="stylesheet" type="text/css" href="${ctx}/static/webuploader/webuploader.css" />
<link rel="stylesheet" type="text/css" href="${ctx}/static/webuploader/uploader_demo.css">
</head>

<body data-color="grey" class="flat">
	<div id="wrapper">
		<div id="header">
			<h1>
				<a href="#">Unicorn Admin</a>
			</h1>
			<a id="menu-trigger" href="#"><i class="fa fa-align-justify"></i></a>
		</div>

		<div id="sidebar">
			<div id="search">
				<input type="text" placeholder="Search here..."/><button type="submit" class="tip-right" title="Search"><i class="fa fa-search"></i></button>
			</div>
			<ul>
				<li class="active"><a href="${ctx}/autocode"><i class="fa fa-home"></i> <span>代码生成器</span></a></li>
			</ul>
		</div>

		<div id="content">
			<div id="content-header">
				<h1>代码生成器</h1>

			</div>
			<div id="breadcrumb">
				<a href="#" title="Go to Home" class="tip-bottom"><i
					class="fa fa-home"></i> 首页</a> <a href="#" class="current">代码生成器</a>
			</div>
	
			<div class="container-fluid">
				<div class="row">
					<div class="col-xs-12">
						<div class="widget-box">
							<div class="widget-title">
								<span class="icon">
									<i class="fa fa-align-justify"></i>									
								</span>
								<h5>录入</h5>
							</div>
							<div class="widget-content nopadding">
								<form id="inputForm"  method="post" class="form-horizontal">
									<div class="form-group">
										<label class="col-sm-3 col-md-3 col-lg-2 control-label">Odoo项目URL</label>
										<div class="col-sm-9 col-md-9 col-lg-10">
											<input name="odooUrl" type="text" class="form-control input-sm" value="http://101.200.124.206:4713"/>
										</div>
									</div>
									
									<div class="form-group">
										<label for="" class="col-sm-3 col-md-3 col-lg-2 control-label">Odoo账号及密码</label>
										<div class="col-sm-9 col-md-9 col-lg-10">
											<div class="row">
												<div class="col-md-6">
													<div class="input-icon icon-sm">
														<i class="fa fa-tint"></i>
														<input name="uname" type='text' class="form-control input-sm" value="admin" />
													</div>
												</div>
												<div class="col-md-6">
													<div class="input-icon on-right icon-sm">
														<input name="pwd" type='text' class="form-control input-sm" value="123456"/>
														<i class="fa fa-laptop"></i>
													</div>
												</div>
											</div>
										</div>
									</div>
									
									<div class="form-group">
										<label class="col-sm-3 col-md-3 col-lg-2 control-label">数据库名</label>
										<div class="col-sm-9 col-md-9 col-lg-10">
											<input name="db" type="text" class="form-control input-sm" value="hospital-saas"/>
										</div>
									</div>
									
									<div class="form-group">
										<label class="col-sm-3 col-md-3 col-lg-2 control-label">模型名</label>
										<div class="col-sm-9 col-md-9 col-lg-10">
											<input name="modelName" type="text" class="form-control input-sm" value="check.list" />
										</div>
									</div>
									
									<div class="form-group">
										<label class="col-sm-3 col-md-3 col-lg-2 control-label">action Id</label>
										<div class="col-sm-9 col-md-9 col-lg-10">
											<input name="action_id" type="text" class="form-control input-sm" value="407"/>
										</div>
									</div>
									
									<div class="form-group">
										<label class="col-sm-3 col-md-3 col-lg-2 control-label">view id</label>
										<div class="col-sm-9 col-md-9 col-lg-10">
											<div class="row">
												<div class="col-md-4">
													<input type="text" name="search_id" class="form-control input-sm" placeholder="输入筛选器视图id" value="" />
													<span class="help-block text-center">search view_id</span>
												</div>
												<div class="col-md-4">
													<input type="text" name="tree_id" class="form-control input-sm" placeholder="输入列表视图id" value="825" />
													<span class="help-block text-center">tree view_id</span>
												</div>
												<div class="col-md-4">
													<input type="text" name="form_id" class="form-control input-sm" placeholder="输入表单视图id" value="824"/>
													<span class="help-block text-center">form view_id</span>
												</div>
											</div>
										</div>
									</div>
									<!-- 
									<div class="form-group" id="webupload-test" >
										<label class="col-sm-3 col-md-3 col-lg-2 control-label">图片上传</label>
										<div class="col-sm-9 col-md-9 col-lg-10">
											<div id="uploader" class="wu-example">
											    <div id="queuelist" class="uploader-list"></div>
											    <div class="btns">
											        <div id="multi"></div>
											        <input type="button" class="btn btn-default" value="上传" id="multiUpload"/> 
											    </div>
											</div>
										</div>
									</div>
									 -->
									 <!-- 
									<div class="form-group" id="webupload-test" >
										<label class="col-sm-3 col-md-3 col-lg-2 control-label">图片上传</label>
										<div class="col-sm-9 col-md-9 col-lg-10">
											<div id="uploader" class="wu-example">
												<div class="queueList">
								                    <div id="dndArea" class="placeholder">
								                        <div id="filePicker"></div>
								                        <p>或将图片拖到这里</p>
								                    </div>
								                    <div class="statusBar" style="display: none;">
									                    <div class="progress">
									                        <span class="text">0%</span>
									                        <span class="percentage"></span>
									                    </div>
									                    <div class="info"></div>
									                    <div class="btns">
									                        <div id="filePicker2"></div>
									                        <div class="uploadBtn">开始上传</div>
									                    </div>
									                </div>
								                </div>
											</div>
										</div>
									</div>
									  -->
									 
									 
									<div class="form-actions">
										<button id="btn-commit" type="button" class="btn btn-primary btn-sm">生成</button> 或 <a id="btn-reset" class="text-danger" href="javascript:void(0);">重置</a>
									</div>
								</form>
							</div>
						</div>
					</div>
				</div>
			</div>
		</div>
		
		
		
		
		
		<div class="row">
			<div id="footer" class="col-xs-12">
					2012 - 2013 &copy; Unicorn Admin. Brought to you by <a href="https://wrapbootstrap.com/user/diablo9983">diablo9983</a>
			</div>
		</div>
	</div>
	
	<%@include file="/WEB-INF/commons/post-include.jsp"%>
	<script type="text/javascript" src="${ctx}/static/webuploader/webuploader.min.js"></script>
	<script src="${ctx}/static/my/js/autocode_main.js"></script>
	<script src="${ctx}/static/my/js/upload2.js"></script>
</body>
</html>
