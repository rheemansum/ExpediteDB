import java.io.File;
import java.io.FileOutputStream;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.Date;
import java.util.ListIterator;
import org.apache.poi.common.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFCreationHelper;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelPackager {
	private ListIterator names_it;
	private ListIterator paths_it;
	private ListIterator folders_it;
	private ListIterator folderpaths_it;
	private ArrayList<Date> dates_ep;
	private XSSFWorkbook masterbook = new XSSFWorkbook();
	
	
	void makesheet(ArrayList names, ArrayList<Date> dates, ArrayList paths, ArrayList folders, ArrayList folderpaths, String sheet_title){
		names_it = names.listIterator();
		paths_it = paths.listIterator();
		folders_it = folders.listIterator();
		folderpaths_it = folderpaths.listIterator();
		this.dates_ep = dates;		
		XSSFSheet new_sheet = masterbook.createSheet(sheet_title);
		XSSFCreationHelper create_helper = masterbook.getCreationHelper();
		XSSFCellStyle cell_style_date = masterbook.createCellStyle();
		cell_style_date.setDataFormat(create_helper.createDataFormat().getFormat("mm/dd/yyyy hh:mm"));        
        int rowid = 1;
        int cellid = 0;               
        Row header_row = new_sheet.createRow(0);
        Cell HeaderCell_1 = header_row.createCell(0);
        HeaderCell_1.setCellValue("Folder");
        Cell HeaderCell_2 = header_row.createCell(1);
        HeaderCell_2.setCellValue("Doc Name");
        Cell HeaderCell_3 = header_row.createCell(2);
        HeaderCell_3.setCellValue("Date");
        int count = 0;
        System.out.println("creating excel sheet...");
        while(names_it.hasNext()){
            XSSFHyperlink file_link = create_helper.createHyperlink(Hyperlink.LINK_FILE);
            XSSFHyperlink folder_link = create_helper.createHyperlink(Hyperlink.LINK_FILE);
            try{
            	String temp_folder_address = folderpaths_it.next().toString();
            	temp_folder_address = URLEncoder.encode(temp_folder_address,"UTF-8");
            	temp_folder_address = temp_folder_address.replace("+","%20");
            	String temp_file_address = paths_it.next().toString();
            	temp_file_address = URLEncoder.encode(temp_file_address,"UTF-8");
            	temp_file_address = temp_file_address.replace("+","%20");
            	file_link.setAddress(temp_file_address);
            	folder_link.setAddress(temp_folder_address);
            }
            catch(Exception ex){
            	ex.printStackTrace();
            }
        	//reset cell to first column
        	cellid = 0;
        	//write data from lists to cells
        	Row new_row = new_sheet.createRow(rowid++);
        	Cell cell_1 = new_row.createCell(cellid++);
        	cell_1.setCellValue(folders_it.next().toString());
        	cell_1.setHyperlink(folder_link);
        	Cell cell_2 = new_row.createCell(cellid++);
        	cell_2.setCellValue(names_it.next().toString()); 
        	cell_2.setHyperlink(file_link);
        	Cell cell_3 = new_row.createCell(cellid++);
        	cell_3.setCellValue(dates_ep.get(count));
        	cell_3.setCellStyle(cell_style_date);
        	count++;
        }
        //autosize column width and apply data filter
        for (int x = 0; x < new_sheet.getRow(0).getPhysicalNumberOfCells();x++){
        	new_sheet.autoSizeColumn(x);
        }
        new_sheet.setAutoFilter(CellRangeAddress.valueOf("A1:C1"));
	}
	
	void closebook(String dest_path){
		try{
        	FileOutputStream fileout = new FileOutputStream(new File(dest_path));
        	masterbook.write(fileout);
        	fileout.close();
	       }
	    catch(Exception ex_output){
	    	ex_output.printStackTrace();
        }  
		System.out.println("complete");		
	}
}
