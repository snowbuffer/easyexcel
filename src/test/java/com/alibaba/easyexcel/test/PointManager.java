package com.alibaba.easyexcel.test;

import java.util.*;
import java.util.stream.Collectors;

/**
 * Description:
 *
 * @author cjb
 * @version V1.0
 * @since 2019-09-20 15:38
 */
public class PointManager {

    private List<Point> pointList = new ArrayList<>();

    private List<List<Object>> titleList;

    public static final String EMPTY = null;

//    public static void main(String[] args) {
//        PointManager manager = new PointManager();
//        manager.addPoint(new Point(2, 2, 3, 4, "@"));
//        manager.addPoint(new Point(5, 6, 2, 5, "%"));
//        manager.addPoint(new Point(7, 3, 2, 6, "&"));
//        manager.calc();
//
//        manager.titleList.forEach(System.out::println);
//
//        manager.pointList.forEach(item -> {
//            System.out.println(item.getPointRangeList());
//        });
//    }

    public void addPoint(Point point) {
        if (point != null) {
            pointList.add(point);
        }
    }

    public void calc() {
        int maxRow = maxRow(pointList);
        int maxColum = maxColumn(pointList);
        drawRectangle(maxRow, maxColum);

        for (Point point : pointList) {
            int rowStartIndex = point.getRowStart() - 1;
            int leftToRightRepeat = point.getLeftToRightRepeat();
            int columStartIndex = point.getColumStart() - 1;
            int upToDownRepeat = point.getUpToDownRepeat();
            populateRectangleForRow(titleList.get(rowStartIndex), columStartIndex, leftToRightRepeat, point.getHeader());
            populateRectangleForColumn(titleList, rowStartIndex, columStartIndex, upToDownRepeat, leftToRightRepeat, point.getHeader());
        }
    }

    private void populateRectangleForColumn(List<List<Object>> titleList, int rowStartIndex, int columStartIndex, int upToDownRepeat, int leftToRightRepeat, Object header) {
        for (int i = 0; i < upToDownRepeat; i++) {
            List<Object> rowList = titleList.get(rowStartIndex + i);
            populateRectangleForRow(rowList, columStartIndex, leftToRightRepeat, header);
        }
    }

    public void populateRectangleForRow(List<Object> list, int columStartIndex, int leftToRightRepeat, Object header) {
        for (int i = 0; i < leftToRightRepeat; i++) {
            int location = columStartIndex + i;

            Object o = list.get(location);
            if (o != null && o.equals(header)) {
                break;
            }
            list.set(columStartIndex + i, header);
            titleList.forEach(System.out::println);
            System.out.println();
        }
    }

    private void drawRectangle(int maxRow, int maxColum) {
        System.out.println(String.format("maxRow = %s, maxColum = %s", maxRow, maxColum));

        titleList = new ArrayList<>();
        for (int i = 0; i < maxRow; i++) {
            List<Object> rowList = new ArrayList<>();
            for (int j = 0; j < maxColum; j++) {
                rowList.add(EMPTY);
            }
            titleList.add(rowList);
        }

        titleList.forEach(System.out::println);
        System.out.println();
    }

    private int maxColumn(List<Point> pointList) {
        Optional<Integer> first = pointList.stream().map((point) -> {
            int columnStart = point.getColumStart();
            int leftToRightRepeat = point.getLeftToRightRepeat();
            int total = columnStart + leftToRightRepeat;
            return total - 1; // leftToRightRepeat 当前左右方向上重复次数，本身就包含起点位置数据
        }).peek((item) -> {
            System.out.println("maxColumn : " + item);
        }).sorted((o1, o2) -> -o1.compareTo(o2)).findFirst();
        if (false == first.isPresent()) {
            throw new RuntimeException("找不到最大列");
        }
        return first.get();
    }

    private int maxRow(List<Point> pointList) {
        Optional<Integer> first = pointList.stream().map((point) -> {
            int rowStart = point.getRowStart();
            int upToDownRepeat = point.getUpToDownRepeat();
            int total = rowStart + upToDownRepeat;
            return total - 1; // upToDownRepeat 当前上下方向上重复次数，本身就包含起点位置数据
        }).peek((item) -> {
            System.out.println("maxRow : " + item);
        }).sorted((o1, o2) -> -o1.compareTo(o2)).findFirst();
        if (false == first.isPresent()) {
            throw new RuntimeException("找不到最大行");
        }
        return first.get();
    }

    public List<List<Object>> getTitleList() {
        return titleList;
    }

    public List<Point> getPointList() {
        return pointList;
    }
}

