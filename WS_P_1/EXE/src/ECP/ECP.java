package ECP;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ECP {
	public static void main(String[] args) throws IOException {
		
		// Create XML File
		File f = new File("C:\\Puneet\\P\\Softwares\\Desktop Ticker\\ECP.xml");
		f.createNewFile();
		
		// Write XML File Headers
		FileWriter fw = new FileWriter(f);
		fw.write("<?xml version=\"1.0\" encoding=\"iso-8859-1\"?>");
		fw.append('\n');
		fw.write("<rss version=\"2.0\" xmlns:content = \"http://purl.org/rss/1.0/modules/content/\">");
		fw.append('\n');
		fw.write("<channel>");
		fw.append('\n');
		fw.append('\n');
		
		// Create Connection to EXECP
		FileInputStream fis = new FileInputStream("C:\\Puneet\\P\\Softwares\\Desktop Ticker\\ECP.xlsx");
		Workbook wb = WorkbookFactory.create(fis);
		org.apache.poi.ss.usermodel.Sheet s1 = wb.getSheet("1");		
		
		// Read EXECP and Write to XML File
		for (int row=2;row<=648;row++) {
		String celldata = s1.getRow(row).getCell(2).getStringCellValue();
		fw.write("<item>");
		fw.append('\n');
		fw.write("<title>"+celldata+"</title>");
		fw.append('\n');
		fw.write("<description>");
		fw.append('\n');
		fw.write("<![CDATA[");
		fw.append('\n');
		fw.write("]]>");
		fw.append('\n');
		fw.write("</description>");
		fw.append('\n');
		fw.write("</item>");
		fw.append('\n');
		fw.append('\n');
		}
		fis.close();
		
		// Write XML File Closing Lines
		fw.write("</channel>");
		fw.append('\n');
		fw.write("</rss>");
		fw.append('\n');
		fw.append('\n');		
		fw.close();	
	}
}
