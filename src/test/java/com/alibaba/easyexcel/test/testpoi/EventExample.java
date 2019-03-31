
/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

package com.alibaba.easyexcel.test.testpoi;

import org.apache.poi.hssf.eventusermodel.HSSFEventFactory;
import org.apache.poi.hssf.eventusermodel.HSSFListener;
import org.apache.poi.hssf.eventusermodel.HSSFRequest;
import org.apache.poi.hssf.record.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

/**
 * This example shows how to use the event API for reading a file.
 */
public class EventExample
        implements HSSFListener {
    private SSTRecord sstrec;

    private int count;

    /**
     * This method listens for incoming records and handles them as required.
     *
     * @param record The record that was found while reading.
     */
    @Override
    public void processRecord(Record record) {
        switch (record.getSid()) {
            // the BOFRecord can represent either the beginning of a sheet or the workbook
            case BOFRecord.sid: // 文件开始
                System.out.println(++count + "=====================BOFRecord.sid start=====================");
                BOFRecord bof = (BOFRecord) record;
                if (bof.getType() == BOFRecord.TYPE_WORKBOOK) {
                    System.out.println("Encountered workbook");
                    // assigned to the class level member
                } else if (bof.getType() == BOFRecord.TYPE_WORKSHEET) {
                    System.out.println("Encountered sheet reference");
                }
                System.out.println("=====================BOFRecord.sid end=====================");
                break;
            case BoundSheetRecord.sid: // 哪个sheet
                System.out.println(++count + "=====================BoundSheetRecord.sid start=====================");
                BoundSheetRecord bsr = (BoundSheetRecord) record;
                System.out.println("New sheet named: " + bsr.getSheetname());
                System.out.println("=====================BoundSheetRecord.sid end=====================");
                break;
            case RowRecord.sid:  // 哪行
                System.out.println(++count + "=====================RowRecord.sid start=====================");
                RowRecord rowrec = (RowRecord) record;
                System.out.println("Row found, first column at "
                        + rowrec.getFirstCol() + " last column at " + rowrec.getLastCol()); // 表格中最大的列
                System.out.println("=====================RowRecord.sid end=====================");
                break;
            case NumberRecord.sid: // 数值单元格
                System.out.println(++count + "=====================NumberRecord.sid start=====================");
                NumberRecord numrec = (NumberRecord) record;
                System.out.println("Cell found with value " + numrec.getValue()
                        + " at row " + numrec.getRow() + " and column " + numrec.getColumn());
                System.out.println("=====================NumberRecord.sid end=====================");
                break;
            // SSTRecords store a array of unique strings used in Excel.
            case SSTRecord.sid: // 表格中出现的唯一的字符串(去重)  ？ 事件机制，为什么可以读取全局的字符串  应该是全局字符串是单独作为常量池存储的
                System.out.println(++count + "=====================SSTRecord.sid start=====================");
                sstrec = (SSTRecord) record;
                for (int k = 0; k < sstrec.getNumUniqueStrings(); k++) {
                    System.out.println("String table value " + k + " = " + sstrec.getString(k));
                }
                System.out.println("=====================SSTRecord.sid end=====================");
                break;
            case LabelSSTRecord.sid:  // 字符串单元格，从左往右，从上往下
                System.out.println(++count + "=====================LabelSSTRecord.sid start=====================");
                LabelSSTRecord lrec = (LabelSSTRecord) record;
                System.out.println("String cell found with value "
                        + sstrec.getString(lrec.getSSTIndex()));
                System.out.println("=====================LabelSSTRecord.sid end=====================");
                break;
        }
    }

    /**
     * Read an excel file and spit out what we find.
     *
     * @param args Expect one argument that is the file to read.
     * @throws IOException When there is an error processing the file.
     */
    public static void main(String[] args) throws IOException {
        // create a new file input stream with the input file specified
        // at the command line
        FileInputStream fin = new FileInputStream("C:\\Users\\admin\\Desktop\\ttttt\\学生账号导入新用户 - 副本 (2).xls");
        // create a new org.apache.poi.poifs.filesystem.Filesystem
        POIFSFileSystem poifs = new POIFSFileSystem(fin);
        // get the Workbook (excel part) stream in a InputStream
        InputStream din = poifs.createDocumentInputStream("Workbook");
        // construct out HSSFRequest object
        HSSFRequest req = new HSSFRequest();
        // lazy listen for ALL records with the listener shown above
        req.addListenerForAllRecords(new EventExample());
        // create our event factory
        HSSFEventFactory factory = new HSSFEventFactory();
        // process our events based on the document input stream
        factory.processEvents(req, din);
        // once all the events are processed close our file input stream
        fin.close();
        // and our document input stream (don't want to leak these!)
        din.close();
        poifs.close();
        System.out.println("done.");
    }
}
