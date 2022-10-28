import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.*;  

public class generateTickets {
    // public static String allSongsListExcelPath =  "C:\\Users\\laiyer\\Desktop\\Tambola_songs.xlsx";
    // public static String allParticipantNamesListExcelPath = "C:\\Users\\laiyer\\Desktop\\namess.xlsx";
    // public static String folderToSavetheTickets = "C:\\Users\\laiyer\\Desktop\\Tickets1\\";
    
    public static String allSongsListExcelPath =  "C:/Users/laiyer/OneDrive - Microsoft/Desktop/Tambola2022/LetsPlayTambolaSongs.xlsx";
    public static String allParticipantNamesListExcelPath = "C:/Users/laiyer/OneDrive - Microsoft/Desktop/Tambola2022/namess1.xlsx";
    public static String folderToSavetheTickets = "C:/Users/laiyer/OneDrive - Microsoft/Desktop/Tambola2022/TicketsNew/";

    public static HashMap<Integer,String> hm_songs = createHashMapSongs(0,0,1);
	public static HashMap<String,Integer> hm_names = createHashMapNames();
    public static List<String> filenames = createFileNamesList();

    public static void main(String args[])
    {
        System.out.println("Total Tickets: "+totalTickets());
            List<String>files6= new ArrayList<String>();
            for(int i=0;i<filenames.size();i++)
            {
                files6.add(filenames.get(i));
                if((i+1)%6==0)
                {
                    System.out.println(); // for test purpose in the terminal output
                    System.out.println("Calling Main"); // for test purpose in the terminal output
                    App1.main(files6);
                    files6.clear();
                }
            }
            System.out.println(); // for test purpose in the terminal output
            System.out.println("Calling Main"); // for test purpose in the terminal output
            App1.main(files6);
    }

    public static List<String> createFileNamesList()
    {
        List<String> filenames = new ArrayList<String>();
         for(String hmvalue : generateTickets.hm_names.keySet())
		 	{
		 		for(int a = 0;a<generateTickets.hm_names.get(hmvalue);a++)
		 		{
                     int t = a+1;
                     String[] n = hmvalue.split("@");
                     String temp;
                     if(t==1)
                     {
                         temp=n[0];
                     }
		 			 else
                      {
                          temp = n[0] + "_"+t;
                      }
                     filenames.add(temp);
                    //filenames.add(hmvalue);
		 		}
		 	}
             System.out.println(filenames);
             return filenames;
    }
public static HashMap<Integer,String> createHashMapSongs(int startRow,int keyCol,int valCol)
{
    HashMap<Integer,String> hm = new HashMap<Integer,String>();
    String path = allSongsListExcelPath;
    try {
        FileInputStream fis = new FileInputStream(path);
        XSSFWorkbook wb = new XSSFWorkbook(fis);
        XSSFSheet sheet1 = wb.getSheet("Sheet1"); 
        int n=startRow; 
        for(Row row: sheet1)     //iteration over row using for each loop  
        {  
            int hmkey = (int)sheet1.getRow(n).getCell(keyCol).getNumericCellValue();
            String hmvalue = sheet1.getRow(n).getCell(valCol).getStringCellValue();
            n++; 
            hm.put(hmkey,hmvalue);
        }
    } catch (IOException e) {
        e.printStackTrace();
    }   
    return hm;
}

public static HashMap<String,Integer> createHashMapNames()
{
    HashMap<String,Integer> hm = new HashMap<String,Integer>();
    String path = allParticipantNamesListExcelPath; 
    try {
         
        //Create an object of FileInputStream class to read excel file
        FileInputStream fis = new FileInputStream(path);
         
        //Create object of XSSFWorkbook class
        XSSFWorkbook wb = new XSSFWorkbook(fis);
         
        //Read excel sheet by sheet name 
        XSSFSheet sheet1 = wb.getSheet("Sheet1"); 
        int n=0; 
        for(Row row: sheet1)     //iteration over row using for each loop  
        {  
            String hmkey = sheet1.getRow(n).getCell(0).getStringCellValue();
            Integer hmvalue=1;
            if(sheet1.getRow(n).getCell(1).getStringCellValue()!="")
            {
                hmvalue = (Integer.parseInt(sheet1.getRow(n).getCell(1).getStringCellValue()))/500;
            }
            n++;
            hm.put(hmkey,hmvalue);
        }
    } catch (IOException e) {
        e.printStackTrace();
    }   
    return hm;
}

// use if required
public static int totalTickets()
	{
		int t = 0;
		for(String hmvalue : generateTickets.hm_names.keySet())
			{
				for(int a = 0;a<generateTickets.hm_names.get(hmvalue);a++)
				{
					t++;
				}
			}
			//System.out.println(t);
		return t;
	}
    public static void printToExcel(String A[][],String filename)
    {  
        try {
			
            XSSFWorkbook workbook = new XSSFWorkbook();
  
            // spreadsheet object
            //XSSFSheet spreadsheet= workbook.createSheet(" Ticket ");

			//XSSFSheet sheet1 = workbook.getSheet("Sheet1");
            XSSFSheet sheet1 = workbook.createSheet("Sheet1");
            sheet1.setDefaultColumnWidth(14); 
    
            // creating a row object
            XSSFRow row;
			//int rowid = sheet1.getLastRowNum();
			//CellStyle headerStyle = workbook.createCellStyle();
			//headerStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
			//headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

			XSSFFont font = ((XSSFWorkbook) workbook).createFont();
			font.setFontName("Arial");
			font.setFontHeightInPoints((short) 10);
			font.setBold(true);

			CellStyle style = workbook.createCellStyle();
			style.setFont(font);
			style.setBorderTop(BorderStyle.MEDIUM);
			style.setBorderBottom(BorderStyle.MEDIUM);
			style.setBorderRight(BorderStyle.MEDIUM);
			style.setBorderLeft(BorderStyle.MEDIUM);

			style.setWrapText(true); 

			style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            
			//style.setFillForegroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex());
            style.setFillForegroundColor(IndexedColors.LIGHT_TURQUOISE.getIndex());

			//style.setLocked(true);
			
            int rowid = 0;
            for(int i=0;i<3;i++)
            {
                int cellid = 0;
                row = sheet1.createRow(rowid++);
                for(int j=0;j<9;j++)
                {
                    Cell cell = row.createCell(cellid++);
                    cell.setCellValue(A[i][j]);
					cell.setCellStyle(style);
                    //System.out.println(A[i][j]);
					//sheet1.getRow(rowid).getCell(cellid).setCellValue(A[i][j]);
                }
				//sheet1.setRowHeight(i,50);
            }
			//  for (int i=0; i<10; i++){
			//  	sheet1.autoSizeColumn(i);
			//   }
            FileOutputStream out = new FileOutputStream(new File(folderToSavetheTickets+filename+".xlsx"));
  
            workbook.write(out);
            out.close();
        } catch (Exception e) {
            e.printStackTrace();
        }   
    }
}
