import java.util.ArrayList;
import java.util.Date;

public class ExpeditingDB_Builder {
	public static void main(String[] args){
		ArrayList transmittals_names = new ArrayList();
		ArrayList<Date> transmittals_dates = new ArrayList();
		ArrayList transmittals_paths = new ArrayList();
		ArrayList transmittals_folders = new ArrayList();
		ArrayList transmittals_folderpaths = new ArrayList();
		ArrayList sitedownloads_names = new ArrayList();
		ArrayList<Date> sitedownloads_dates = new ArrayList();
		ArrayList sitedownloads_paths = new ArrayList();
		ArrayList sitedownloads_folders = new ArrayList();
		ArrayList sitedownloads_folderpaths = new ArrayList();
		String transmittals_src = "C:\\Transmittals";
		String sitedownloads_src = "C:\\SiteDownlods";
		String destination = "C:\\Folder";
		FileScrape transmittals = new FileScrape();
		FileScrape sitedownloads = new FileScrape();
		
		transmittals.scrape(transmittals_src);
		transmittals_names = transmittals.getnameslist();
		transmittals_dates = transmittals.getdateslist();
		transmittals_paths = transmittals.getpathslist();
		transmittals_folders = transmittals.getfolderslist();
		transmittals_folderpaths = transmittals.getfolderpathslist();
		
		sitedownloads.scrape(sitedownloads_src);
		sitedownloads_names = sitedownloads.getnameslist();
		sitedownloads_dates = sitedownloads.getdateslist();
		sitedownloads_paths = sitedownloads.getpathslist();
		sitedownloads_folders = sitedownloads.getfolderslist();
		sitedownloads_folderpaths = sitedownloads.getfolderpathslist();
		
		ExcelPackager master = new ExcelPackager();
		master.makesheet(transmittals_names, transmittals_dates, transmittals_paths, transmittals_folders, transmittals_folderpaths, "Transmittals");
		master.makesheet(sitedownloads_names, sitedownloads_dates, sitedownloads_paths, sitedownloads_folders, sitedownloads_folderpaths, "SiteDownloads");
		master.closebook(destination);
	}
}