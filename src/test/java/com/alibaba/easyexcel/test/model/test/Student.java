package com.alibaba.easyexcel.test.model.test;

import com.alibaba.easyexcel.test.model.test.converter.DefaultStringConverter;
import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.metadata.RowErrorModel;

/**
 * Description:
 *
 * @author cjb
 * @version V1.0
 * @since 2019-03-30 15:12
 */
public class Student extends RowErrorModel {

    @ExcelProperty(index = 0, value = {"学生姓名"} )
    private String name;

    @ExcelProperty(index = 1, value = {"所在学校"}, convertor = DefaultStringConverter.class)
    private String school;

    @ExcelProperty(index = 2, value = {"学籍号"})
    private String grade;

    @ExcelProperty(index = 3, value = {"年级"})
    private String nianji;

    @ExcelProperty(index = 4, value = {"班级"})
    private String classRoom;

    @ExcelProperty(index = 5, value = {"性别"})
    private String gender;

    @ExcelProperty(index = 6, value = {"家长姓名"})
    private String masterName;

    @ExcelProperty(index = 7, value = {"家长联系方式"})
    private Long telephone;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getSchool() {
        return school;
    }

    public void setSchool(String school) {
        this.school = school;
    }

    public String getGrade() {
        return grade;
    }

    public void setGrade(String grade) {
        this.grade = grade;
    }

    public String getNianji() {
        return nianji;
    }

    public void setNianji(String nianji) {
        this.nianji = nianji;
    }

    public String getClassRoom() {
        return classRoom;
    }

    public void setClassRoom(String classRoom) {
        this.classRoom = classRoom;
    }

    public String getGender() {
        return gender;
    }

    public void setGender(String gender) {
        this.gender = gender;
    }

    public String getMasterName() {
        return masterName;
    }

    public void setMasterName(String masterName) {
        this.masterName = masterName;
    }

    public Long getTelephone() {
        return telephone;
    }

    public void setTelephone(Long telephone) {
        this.telephone = telephone;
    }

    @Override
    public String toString() {
        final StringBuffer sb = new StringBuffer("Student [");
        sb.append("name=").append(name);
        sb.append(", school=").append(school);
        sb.append(", grade=").append(grade);
        sb.append(", nianji=").append(nianji);
        sb.append(", classRoom=").append(classRoom);
        sb.append(", gender=").append(gender);
        sb.append(", masterName=").append(masterName);
        sb.append(", telephone=").append(telephone);
        sb.append(", RowErrorModel=").append(super.toString());
        sb.append("]");
        return sb.toString();
    }
}

