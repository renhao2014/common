package com.template.util.model;

import java.io.Serializable;
import java.util.Date;


/**
 * 模板-实体
 * @author Renhao
 * @version 1.0
 */

public class Template implements Serializable{

    private static final long serialVersionUID = -8893494706131531685L;

    private String templateId;
    //String字段
    private String templateString;
    //整数字段
    private Integer templateInt;
    //小数字段
    private Double templateDouble;
    //时间字段
    private Date templateDate;

    public static long getSerialVersionUID() {
        return serialVersionUID;
    }

    public String getTemplateId() {
        return templateId;
    }

    public void setTemplateId(String templateId) {
        this.templateId = templateId;
    }

    public String getTemplateString() {
        return templateString;
    }

    public void setTemplateString(String templateString) {
        this.templateString = templateString;
    }

    public Integer getTemplateInt() {
        return templateInt;
    }

    public void setTemplateInt(Integer templateInt) {
        this.templateInt = templateInt;
    }

    public Double getTemplateDouble() {
        return templateDouble;
    }

    public void setTemplateDouble(Double templateDouble) {
        this.templateDouble = templateDouble;
    }

    public Date getTemplateDate() {
        return templateDate;
    }

    public void setTemplateDate(Date templateDate) {
        this.templateDate = templateDate;
    }
}
