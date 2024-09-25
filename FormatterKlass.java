
package groupID.formatpenaltyfile;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import static java.lang.Thread.sleep;
import java.util.Scanner;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import org.apache.poi.xssf.usermodel.XSSFSheet; 
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;

/**
 *
 * @author csomabalint
 */
public class FormatterKlass {
    
    public static void main(String[] args) throws FileNotFoundException {

        File directoryPath = new File("/Users/csomabalint/Java/Teszt/DLH Penalty Reports CSV/");
        //File directoryPath = new File("D:/Apps/PENALTY/run_abc/");
        String contents[] = directoryPath.list(); //List of all files and directories
        Pattern p = null;
        for (int k = 0; k < 2; k++) {
            if (k == 0){
                p = Pattern.compile("Penalty Summary Parameter_.*csv");
            }if (k == 1){
                p = Pattern.compile("Penalty Summary Parameter Internal.*csv");
            }
            
String content = "";
for (int i = 0; i < contents.length; i++) {
            content = contents[i];                   
            Matcher m = p.matcher(content);
            boolean b = m.matches();
            if(b == true){
                break;
            }
            else {
                continue;
            }
        }
String filePathIn = "/Users/csomabalint/Java/Teszt/DLH Penalty Reports CSV/" + content;
//String filePathIn = "D:/Apps/PENALTY/run_abc/" + content;
String filePathOut = filePathIn.substring(0, filePathIn.length()-3) + "xlsx";
FileInputStream input = new FileInputStream(filePathIn);
Scanner scan = new Scanner(input,"CP1252");
String[] something = scan.nextLine().split(";");
int colToFormat = 0;
        for (int i = 0; i < something.length; i++) {
            String string = something[i];
            if(string != null){
            if (string.equalsIgnoreCase("Calc. Number of Deviations"))
                colToFormat = i;
            }
            else {
                System.out.println("Column name to format changed!");
    try {
        sleep(10000);
    } catch (InterruptedException ex) {
        Logger.getLogger(FormatterKlass.class.getName()).log(Level.SEVERE, null, ex);
    }
            System.exit(1);
            }
        }
String[] temp = scan.nextLine().split(";");
int colNumber = temp.length;
scan.close();

XSSFWorkbook w = new XSSFWorkbook();   //above Excel 2003 
XSSFSheet sheet = w.createSheet("sheet");
XSSFCellStyle style = w.createCellStyle();
XSSFCellStyle style2 = w.createCellStyle();
XSSFDataFormat df = w.createDataFormat();
XSSFDataFormat df2 = w.createDataFormat();
style.setDataFormat(df.getFormat("@"));
style2.setDataFormat(df2.getFormat("###,###,###,###,###,##0.00 €"));
w.getSheet("sheet").getColumnHelper().setColBestFit(colToFormat, true);
w.getSheet("sheet").getColumnHelper().setColDefaultStyle(colToFormat, style);

Scanner sca = new Scanner(new FileInputStream(filePathIn),"CP1252");
int rowNumber = 0;
while(sca.hasNextLine()){
   if(sca.nextLine() != "")
   rowNumber++;
}
sca.close();
FileInputStream input2 = new FileInputStream(filePathIn); 
Scanner scan2 = new Scanner(input2,"CP1252");
        for (int i = 0; i < rowNumber; i++) {
            String[] row = scan2.nextLine().split(";");
            XSSFRow row2 = sheet.createRow(i);
            for (int j = 0; j < colNumber; j++) {
                try {
                    row2.createCell(j).setCellValue(Double.parseDouble(row[j]));
                    //row2.getCell(j).setCellStyle(style2);//.replace("Ä","€")
                    } catch(NumberFormatException ex) {
                        try {
                        if(i>0){
                        if(j==8 || j==11 || j==12){
                            row2.createCell(j).setCellValue(Double.parseDouble(row[j].replace(",","")));
                           row2.getCell(j).setCellStyle(style2);
                           continue;
                        } 
                        }
                    } catch(NumberFormatException ex2) {
                              //Logger.getLogger(FormatterKlass.class.getName()).log(Level.SEVERE, null, ex2);
                              row2.createCell(j).setCellValue(row[j].replace("(€)",""));
                                   }
                    row2.createCell(j).setCellValue(row[j]);
                    } 
               
            }  
            
        }
        for (int i = 0; i < colNumber; i++) {
            w.getSheet("sheet").autoSizeColumn(i);
            
        }
        try {
FileOutputStream output = new FileOutputStream(filePathOut);
w.write(output);
        } catch (IOException ex) {
            Logger.getLogger(FormatterKlass.class.getName()).log(Level.SEVERE, null, ex);
        }
       }
    PenaltySummary ps = new PenaltySummary();
    ps.penaltySummaryFormatter();
    }
}


    
    

