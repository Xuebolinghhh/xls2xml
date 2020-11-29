package com.xls2xml;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.UnsupportedEncodingException;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;


import org.dom4j.Document;
import org.dom4j.Element;
import org.dom4j.DocumentHelper;
import org.dom4j.io.OutputFormat;
import org.dom4j.io.XMLWriter;

public class Version2 {

	public static void main(String[] args) throws Exception {
		Version2 file = new Version2();
		file.xls2xml("test.xls");

	}
	
	public void xls2xml(String excelPath) throws Exception {
		
		Workbook readwb=null;
		
		try {
			
			readwb = Workbook.getWorkbook(new File(excelPath));
			
			// using Dom4j package to build a doc
			Document doc = DocumentHelper.createDocument();
			
			// add a root element into the doc
			Element root = doc.addElement("root");
			
			for(int m = 0; m<readwb.getNumberOfSheets();m++) {
				Sheet sheet = readwb.getSheet(m);
				
				// get the row and the column number of excel
				int nColumns = sheet.getColumns();
				int nRows = sheet.getRows();
				
				//get the title of each column
				Cell[] firstCells = sheet.getRow(0);
				
				// loop each row
				for (int i=1; i<nRows; i++) {
					Element rowElement =root.addElement("rowElement"+i);
					
					for(int j=0; j<nColumns; j++) {
						Cell cell = sheet.getCell(j, i);
						if(cell.getContents() == "") {
							continue;
						}
						Element columnElement = rowElement.addElement(
							firstCells[j].getContents());
						columnElement.addText(cell.getContents());					
					}
				}
				
			}
			
			// define xml format
			OutputFormat format = new OutputFormat();
			format.setIndentSize(2);  // è¡Œç¼©è¿›
            format.setNewlines(true); // ä¸€ä¸ªç»“ç‚¹ä¸ºä¸€è¡Œ
            format.setTrimText(true); // åŽ»é‡�ç©ºæ ¼
            format.setPadText(true);
            format.setNewLineAfterDeclaration(false);
            
         // è¾“å‡ºxmlæ–‡ä»¶
            XMLWriter writer = new XMLWriter(new FileOutputStream("test1.xml"), format);
            writer.write(doc);
            System.out.println("dom4j CreateDom4j success!");
		} catch (UnsupportedEncodingException e) {
			e.printStackTrace();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			readwb.close();
		}
		
		
	}

}
