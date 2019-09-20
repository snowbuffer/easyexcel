package com.alibaba.easyexcel.test;

import com.alibaba.excel.util.StringUtils;

import java.util.ArrayList;
import java.util.List;

/**
 * Description:
 *
 * @author cjb
 * @version V1.0
 * @since 2019-09-20 14:59
 */
public class Point {
    private int rowStart; // 行号，从1 开始
    private int columStart; // 列号，从1 开始
    private int leftToRightRepeat; // 左右方向重复次数 包含point坐标
    private int upToDownRepeat; // 上下向上重复次数 包含point坐标
    private Object header;


    // rowStart + upToDownRepeat => 行数
    // columStart + leftToRightRepeat => 列数
    public Point(int rowStart, int columnStart, int leftToRightRepeat, int upToDownRepeat, Object header) {
        check(rowStart, columnStart, leftToRightRepeat, upToDownRepeat, header);
        this.rowStart = rowStart;
        this.columStart = columnStart;
        this.leftToRightRepeat = leftToRightRepeat;
        this.upToDownRepeat = upToDownRepeat;
        this.header = header;
    }

    public List<Integer> getPointRangeList() {
        List<Integer> rs = new ArrayList<>();
        // 起始行
        rs.add(rowStart - 1);
        // 结束行
        rs.add(rowStart + upToDownRepeat - 2);
        // 起始列
        rs.add(columStart -1);
        // 结束列
        rs.add(columStart + leftToRightRepeat -2);
        return rs;
    }

    public int getRowStart() {
        return rowStart;
    }

    public int getColumStart() {
        return columStart;
    }

    public int getLeftToRightRepeat() {
        return leftToRightRepeat;
    }

    public int getUpToDownRepeat() {
        return upToDownRepeat;
    }

    public Object getHeader() {
        return header;
    }

    private void check(int rowStart, int columStart, int rowDataRepeat, int columDataRepeat, Object header) {
        if (rowStart <= 0) {
            throw new RuntimeException("rowStart 不能小于 0");
        }
        if (columStart <= 0) {
            throw new RuntimeException("columStart 不能小于 0");
        }
        if (rowDataRepeat <= 0) {
            throw new RuntimeException("rowDataRepeat 不能小于 0");
        }
        if (columDataRepeat <= 0) {
            throw new RuntimeException("columDataRepeat 不能小于 0");
        }
        if (StringUtils.isEmpty(header)) {
            throw new RuntimeException("header 不能为empty");
        }
    }

    @Override
    public String toString() {
        final StringBuffer sb = new StringBuffer("Point [");
        sb.append("rowStart=").append(rowStart);
        sb.append(", columStart=").append(columStart);
        sb.append(", leftToRightRepeat=").append(leftToRightRepeat);
        sb.append(", upToDownRepeat=").append(upToDownRepeat);
        sb.append(", header=").append(header);
        sb.append("]");
        return sb.toString();
    }
}

