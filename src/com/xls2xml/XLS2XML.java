package com.xls2xml;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;

import java.io.File;
import java.io.FileOutputStream;

import org.jdom.Document;
import org.jdom.Element;
import org.jdom.output.Format;
import org.jdom.output.XMLOutputter;

public class XLS2XML {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		XLS2XML file = new XLS2XML();
		file.excelConvertXml("test.xls", "test.xml");
	}
	
	/**
	 * convert excel to xml
	 * 
	 * @param excelPath
	 *  	the path of excel file
	 * 
	 * @param xmlPath
	 * 		the path of xml file
	 */
	
	public void excelConvertXml(String excelPath, String xmlPath) {
		
		Workbook readwb = null;
		
		try {
			
			readwb = Workbook.getWorkbook(new File(excelPath));
			
			// build a root element
			Element root = new Element("root");
			
			// add the root element into the Doc
			Document doc = new Document(root);
			
			// loop each sheet 
			for(int m = 0; m<readwb.getNumberOfSheets();m++) {
				Sheet sheet = readwb.getSheet(m);
				
				// get the row and the column number of excel
				int nColumns = sheet.getColumns();
				int nRows = sheet.getRows();
				
				//get the title of each column
				Cell[] firstCells = sheet.getRow(0);
				
				// loop each row
				for (int i=1; i<nRows; i++) {
					Element row = new Element("dataDetail");
					
					for(int j=0; j<nColumns; j++) {
						Cell cell = sheet.getCell(j, i);
						if(cell.getContents() == "") {
							continue;
						}
						Element column = new Element(firstCells[j].getContents());
						column.setText(cell.getContents());
						row.addContent(column);
					}
				root.addContent(row);
				}
			}
			// format the content
			Format format = Format.getPrettyFormat();
			XMLOutputter XMLOut = new XMLOutputter(format);
			XMLOut.output(doc, new FileOutputStream(xmlPath));
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			readwb.close();
			System.out.println("run over");
		}
		
	}
}
