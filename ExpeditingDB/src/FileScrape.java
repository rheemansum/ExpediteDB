import java.io.File;
import java.util.ArrayList;
import java.util.Date;

public class FileScrape {
	private ArrayList names_list = new ArrayList();
	private ArrayList<Date> dates_list = new ArrayList<Date>();
	private ArrayList folders_list = new ArrayList();
	private ArrayList paths_list = new ArrayList();
	private ArrayList folderpaths_list = new ArrayList();
	
	void scrape(String x_src){
		int extensionlength = 4;
		File temp_dir = new File(x_src);
		File [] dir_filelist = temp_dir.listFiles();
		for(int x = 0; x < dir_filelist.length; x++){
			String temp_file_name = dir_filelist[x].getName().toLowerCase();
			if(dir_filelist[x].isFile() && temp_file_name.endsWith(".pdf")){
				names_list.add(dir_filelist[x].getName().substring(0, dir_filelist[x].getName().length()-extensionlength));
				Date temp_date = new Date(dir_filelist[x].lastModified());
				dates_list.add(temp_date);
				paths_list.add(dir_filelist[x].getAbsolutePath());
				folders_list.add(temp_dir.getName());
				folderpaths_list.add(temp_dir.getAbsolutePath());
				System.out.println(temp_dir.getName());
				System.out.println(temp_dir.getAbsolutePath());
			}
			else if(dir_filelist[x].isDirectory()){
				String new_source = dir_filelist[x].getPath().toLowerCase();
				System.out.println(dir_filelist[x].getName().toLowerCase());
				scrape(new_source);			
			}
		}
	}
	
	ArrayList getnameslist(){
		return names_list;
	}
	
	ArrayList<Date> getdateslist(){
		return dates_list;
	}
	
	ArrayList getfolderslist(){
		return folders_list;
	}
	
	ArrayList getfolderpathslist(){
		return folderpaths_list;
	}
	
	ArrayList getpathslist(){
		return paths_list;
	}	
}
