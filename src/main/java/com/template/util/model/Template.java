package com.template.util.model;

import java.io.Serializable;
import java.util.Date;


public class Template implements Serializable {
	private static final long serialVersionUID = 1L;
	// 模板主键
	private String id;
	// 创建时间
	private Date createTime;
	// 名称
	private String templateName;
	// 状态
	private Integer templateStatus;
	// 类型
	private Integer templateType;

	public String getId() {
		return id;
	}

	public void setId(String id) {
		this.id = id;
	}

	public void setCreateTime(Date createTime) {
		this.createTime = createTime;
	}

	public Date getCreateTime() {
		return createTime;
	}

	public void setTemplateName(String templateName) {
		this.templateName = templateName;
	}

	public String getTemplateName() {
		return templateName;
	}

	public void setTemplateStatus(Integer templateStatus) {
		this.templateStatus = templateStatus;
	}

	public Integer getTemplateStatus() {
		return templateStatus;
	}

	public void setTemplateType(Integer templateType) {
		this.templateType = templateType;
	}

	public Integer getTemplateType() {
		return templateType;
	}
}
