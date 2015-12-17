import java.io.*;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.Date;
import java.util.ListIterator;
import java.util.Vector;
import org.apache.poi.common.usermodel.Hyperlink;
import org.apache.poi.hssf.util.CellRangeAddress;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFCreationHelper;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.commons.io.FileUtils;
import javax.swing.JOptionPane;
import javax.swing.JFrame;



public class ExpeditingDB_Builder{
	
	@SuppressWarnings("deprecation")
	public static void main(String[] args) throws IOException{
		//setup excel workbook and sheets
		XSSFWorkbook masterbook = new XSSFWorkbook();
        XSSFSheet TBLDocLogSheet = masterbook.createSheet("TBL_DocumentLogs");
        XSSFSheet TransmittalsSheet = masterbook.createSheet("Transmittals");
        XSSFSheet SiteDownloadsSheet = masterbook.createSheet("SiteDownloads");
        XSSFCreationHelper createHelper = masterbook.getCreationHelper();
        XSSFCellStyle CellStyleDate = masterbook.createCellStyle();
        XSSFRow ExcelRow;
        XSSFRow HeaderRow;
        int extensionlength = 4; //length of ".***" 
    	JFrame frame = new JFrame();
    	JOptionPane.showMessageDialog(frame, "starting");
 
        try {
        	//copy and paste access file from network to local drive and open
        	String MDLdestPath = "C:/Projects/DatabaseFolder";
        	String MDLsrcPath = "Q:/ProjectDatabase/TestDatabase.mdb";
        	File destDir = new File(MDLdestPath);
        	File srcFile = new File(MDLsrcPath);
        	FileUtils.copyFileToDirectory(srcFile, destDir);

            Connection conn = DriverManager.getConnection("jdbc:ucanaccess://C:/Projects/DatabaseFolder/TestDatabase.mdb;jackcessOpener=CryptCodecOpener", "", "");
            Statement KarbalaDatabase_st = conn.createStatement();
            ResultSet rs = KarbalaDatabase_st.executeQuery("SELECT * FROM [TBL_Document Logs]");
            ResultSetMetaData metaData = rs.getMetaData();
            //pull data from the TBL_Document Logs database
            int ncol = metaData.getColumnCount();
            Vector<String> columnNames = new Vector<String>();
            for(int column = 0; column < ncol; column++){
            	columnNames.addElement(metaData.getColumnLabel(column+1));
            }
            Vector<Vector> rows = new Vector<Vector>();
            while(rs.next()){
            	Vector newRow = new Vector();
            	for(int i = 1; i <= ncol; i++){
            		newRow.addElement(rs.getObject(i));
            	}
            	rows.addElement(newRow);            	
            }
            
            //write data from database to excel file
            int mdb_rowid = 0;
            int mdb_cellid = 0;
            HeaderRow = TBLDocLogSheet.createRow(mdb_rowid++);
        	for(int z=0;z<columnNames.size();z++){        		
        		Cell celltitle = HeaderRow.createCell(mdb_cellid++);
        		celltitle.setCellValue(columnNames.get(z).toString());
        	}
            
            for(int x=0;x<rows.size();x++){
            	mdb_cellid = 0;
            	ExcelRow = TBLDocLogSheet.createRow(mdb_rowid++);            	
            	for(int y=0;y<rows.get(x).size();y++){
            		Cell cell = ExcelRow.createCell(mdb_cellid++);
            		if (rows.get(x).get(y) == null){
            			cell.setCellValue("");
            		}
            		else{
            			cell.setCellValue(rows.get(x).get(y).toString());          		
            		}
            	}
            }               
        }
        catch (SQLException ex) {
            ex.printStackTrace();
        }
        //autosize column width and add data filter
        for (int x = 0; x < TBLDocLogSheet.getRow(0).getPhysicalNumberOfCells();x++){
        	TBLDocLogSheet.autoSizeColumn(x);
        }
        TBLDocLogSheet.setAutoFilter(CellRangeAddress.valueOf("A1:AN1"));
        
        //Scrape network folders and files
		String TRANSMITTALS_SOURCE = "O:/Engineering_Projects/Project_File_Cabinets/Transmittals w-/Transmittals To";
		String SITEDOWNLOADS_SOURCE = "Q:/Site Downloads (FTP & E-mails)/Documents/ThisProject";

		ArrayList<String> TRANSMITTAL_NAME_LIST = new ArrayList<String>();
		ArrayList<Date> TRANSMITTAL_TIME_LIST = new ArrayList<Date>();
		ArrayList<String> file_link_transmittal_array = new ArrayList<String>();
		
		//Dig through Transmittals folder and pull all transmittal cover page PDFs and modified date into vectors
		File TRANSMITTALS_DIR = new File(TRANSMITTALS_SOURCE);
		File [] RECIPIENTS_fLIST = TRANSMITTALS_DIR.listFiles();
		//System.out.println(Arrays.asList(RECIPIENTS_fLIST));
		for(int i=0; i < RECIPIENTS_fLIST.length; i++){
			if(RECIPIENTS_fLIST[i].isDirectory()){
				File [] ALL_TRANSMITTALS = RECIPIENTS_fLIST[i].listFiles();
				for(int j=0; j < ALL_TRANSMITTALS.length; j++){
					String tempCaseName = ALL_TRANSMITTALS[j].getName().toLowerCase();
					if(ALL_TRANSMITTALS[j].isFile() && tempCaseName.endsWith("pdf")){
						TRANSMITTAL_NAME_LIST.add(ALL_TRANSMITTALS[j].getName().substring(0, ALL_TRANSMITTALS[j].getName().length()-extensionlength));
						Date tempDate = new Date(ALL_TRANSMITTALS[j].lastModified());
						TRANSMITTAL_TIME_LIST.add(tempDate);
						file_link_transmittal_array.add(ALL_TRANSMITTALS[j].getAbsolutePath());
					}
				}	
			}			
		}


        CellStyleDate.setDataFormat(createHelper.createDataFormat().getFormat("mm/dd/yyyy hh:mm"));
        
        int rowid = 1;
        int cellid = 0;
        
        //write vector data into excel
        ListIterator<String> transmittalnameIT = TRANSMITTAL_NAME_LIST.listIterator();
        Row TransmittalsHeaderRow = TransmittalsSheet.createRow(0);
        Cell HeaderCell_1 = TransmittalsHeaderRow.createCell(0);
        HeaderCell_1.setCellValue("Transmittal Name");
        Cell HeaderCell_2 = TransmittalsHeaderRow.createCell(1);
        HeaderCell_2.setCellValue("Date");
        
        int count = 0;
        while(transmittalnameIT.hasNext()){
            XSSFHyperlink file_link_transmittal=createHelper.createHyperlink(Hyperlink.LINK_FILE);
        	String temp_transmittal_file_link_address = file_link_transmittal_array.get(count).toString();
        	temp_transmittal_file_link_address = temp_transmittal_file_link_address.replace("\\", "/");
        	temp_transmittal_file_link_address = temp_transmittal_file_link_address.replace(" ","%20");
        	file_link_transmittal.setAddress(temp_transmittal_file_link_address);
        	
        	cellid = 0;
        	Row TransmittalsRow = TransmittalsSheet.createRow(rowid++);
        	Cell cell_1 = TransmittalsRow.createCell(cellid++);
        	cell_1.setCellValue(transmittalnameIT.next().toString()); 
        	cell_1.setHyperlink(file_link_transmittal);
        	Cell cell_2 = TransmittalsRow.createCell(cellid++);
        	cell_2.setCellValue(TRANSMITTAL_TIME_LIST.get(count));
        	cell_2.setCellStyle(CellStyleDate);
        	count++;
        }
        //autosize column width and apply data filter
        for (int x = 0; x < TransmittalsSheet.getRow(0).getPhysicalNumberOfCells();x++){
        	TransmittalsSheet.autoSizeColumn(x);
        }
        TransmittalsSheet.setAutoFilter(CellRangeAddress.valueOf("A1:C1"));
        
        //Dig through Site Downloads Folder and pull pdf files, folder, and modified time into vectors
		ArrayList<String> SITEDOWNLOAD_NAMEDATE = new ArrayList<String>();
		ArrayList<String> SITEDOWNLOAD_DOCNAME = new ArrayList<String>();
		ArrayList<Date> SITEDOWNLOAD_TIME_LIST = new ArrayList<Date>();
		ArrayList<String> SITEDOWNLOAD_DOCPATH = new ArrayList<String>();
		ArrayList<String> file_link_SD_Array = new ArrayList<String>();
		
		File SITEDOWNLOADS_DIR = new File(SITEDOWNLOADS_SOURCE);
		File [] SITEDOWNLOADS_SENDER_LIST = SITEDOWNLOADS_DIR.listFiles();
		//System.out.println(Arrays.asList(SITEDOWNLOADS_SENDER_LIST));
		for(int i=0; i<SITEDOWNLOADS_SENDER_LIST.length; i++){
			if(SITEDOWNLOADS_SENDER_LIST[i].isDirectory()){
				File[] FOLDER_RECEIVED_ARRAY = SITEDOWNLOADS_SENDER_LIST[i].listFiles();
				for(int j=0; j<FOLDER_RECEIVED_ARRAY.length; j++){
					if(FOLDER_RECEIVED_ARRAY[j].isDirectory()){
						File[] DOCUMENTS_RECEIVED_ARRAY = FOLDER_RECEIVED_ARRAY[j].listFiles();
						for(int k=0; k<DOCUMENTS_RECEIVED_ARRAY.length; k++){
							String tempcaseSD = DOCUMENTS_RECEIVED_ARRAY[k].getName().toLowerCase();
							if(DOCUMENTS_RECEIVED_ARRAY[k].isFile() && tempcaseSD.endsWith("pdf")){
								SITEDOWNLOAD_NAMEDATE.add(FOLDER_RECEIVED_ARRAY[j].getName());
								SITEDOWNLOAD_DOCNAME.add(DOCUMENTS_RECEIVED_ARRAY[k].getName().substring(0, DOCUMENTS_RECEIVED_ARRAY[k].getName().length()-extensionlength));
								SITEDOWNLOAD_DOCPATH.add(DOCUMENTS_RECEIVED_ARRAY[k].getAbsolutePath());
								file_link_SD_Array.add(DOCUMENTS_RECEIVED_ARRAY[k].getAbsolutePath());
								System.out.println(DOCUMENTS_RECEIVED_ARRAY[k].getAbsolutePath());
								Date stempdate = new Date(DOCUMENTS_RECEIVED_ARRAY[k].lastModified());
								SITEDOWNLOAD_TIME_LIST.add(stempdate);							
							}
						}						
					}
				}
			}
		}
		
		
		//
		ListIterator<String> SitedownloadNameDateIT = SITEDOWNLOAD_NAMEDATE.listIterator();
		ListIterator<String> SitedownloadDocNameIT = SITEDOWNLOAD_DOCNAME.listIterator();

        rowid = 1;
        cellid = 0;
        int s_count = 0;
        Row HeaderRow2 = SiteDownloadsSheet.createRow(0);
        Cell HeaderCell_2_1 = HeaderRow2.createCell(0);
        HeaderCell_2_1.setCellValue("Site Download Folder");
        Cell HeaderCell_2_2 = HeaderRow2.createCell(1);
        HeaderCell_2_2.setCellValue("Document Name");
        Cell HeaderCell_2_3 = HeaderRow2.createCell(2);
        HeaderCell_2_3.setCellValue("Date");
        
        while(SitedownloadNameDateIT.hasNext()){
            XSSFHyperlink file_link_SD=createHelper.createHyperlink(Hyperlink.LINK_FILE);
        	String tempfilelinkaddress = file_link_SD_Array.get(s_count).toString();
        	tempfilelinkaddress = tempfilelinkaddress.replace("\\", "/");
        	tempfilelinkaddress = tempfilelinkaddress.replace(" ","%20");
        	file_link_SD.setAddress(tempfilelinkaddress);
        	cellid = 0;
        	Row SiteDonwloadsRow = SiteDownloadsSheet.createRow(rowid++);
        	Cell cell_1 = SiteDonwloadsRow.createCell(cellid++);
        	cell_1.setCellValue(SitedownloadNameDateIT.next().toString());
        	Cell cell_2 = SiteDonwloadsRow.createCell(cellid++);
        	cell_2.setCellValue(SitedownloadDocNameIT.next().toString());
        	cell_2.setHyperlink(file_link_SD);
        	Cell cell_3 = SiteDonwloadsRow.createCell(cellid++);
        	cell_3.setCellValue(SITEDOWNLOAD_TIME_LIST.get(s_count));
        	cell_3.setCellStyle(CellStyleDate);
        	s_count++;
        }
        for (int x = 0; x < SiteDownloadsSheet.getRow(0).getPhysicalNumberOfCells();x++){
        	SiteDownloadsSheet.autoSizeColumn(x);
        }
        SiteDownloadsSheet.setAutoFilter(CellRangeAddress.valueOf("A1:C1"));
        try{
        	String ExcelDestPath = "O:/Engineering_Projects/Project_File_Cabinets/Expediting_DB.xlsx";
        	FileOutputStream out = new FileOutputStream(new File(ExcelDestPath));
        	masterbook.write(out);
        	out.close();
//        	System.out.println("File complete");
        	JOptionPane.showMessageDialog(frame, "complete");

	       }
	    catch(Exception e){
        	e.printStackTrace();
        }  
//		System.out.println("complete");
	}
		
}