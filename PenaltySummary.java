
package groupID.formatpenaltyfile;

//import com.sun.tools.javac.util.ArrayUtils;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Scanner;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;

/**
 *
 * @author csomabalint
 */
class PenaltySummary{

    protected void penaltySummaryFormatter() throws FileNotFoundException {
//    Scanner sca = new Scanner(System.in);
//    System.out.println("Penalty Summary input:");
//    String in = sca.nextLine();
//    System.out.println("Penalty Summary output:");
//    String out = sca.nextLine();
//    sca.close();

File directoryPath = new File("/Users/csomabalint/Java/Teszt/DLH Penalty Reports CSV/");
//File directoryPath = new File("D:/Apps/PENALTY/run_abc/");
String contents[] = directoryPath.list(); //List of all files and directories
Pattern p = Pattern.compile("Penalty Summary_.*csv");
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

Scanner numberOfRows = new Scanner(new FileInputStream(filePathIn),"CP1252");
int rowNumber = 0;
while(numberOfRows.hasNextLine()){
    if(numberOfRows.nextLine() != "")
    rowNumber++;
}
numberOfRows.close();

Scanner sc2 = new Scanner(new FileInputStream(new File(filePathIn)),"CP1252");
String[] colNumberArray = sc2.nextLine().split(";");
Scanner sc = new Scanner(new FileInputStream(new File(filePathIn)),"CP1252");
String[] data = new String[colNumberArray.length];
String[][] datas = new String[rowNumber][colNumberArray.length];
int z =0; //number of rows - 1
for(int i = 0; i<rowNumber;i++){
            data = sc.nextLine().split(";");
            for (int j = 0; j < colNumberArray.length; j++) {
                datas[z][j] = data[j];
                
            }
            z++;
}
        
        
        int colNumber = colNumberArray.length;
        String[] header = datas[0];
        int iPenalty = 0;
        int iEarn_back = 0;
        for (int i = 0; i < header.length; i++) {
            if(header[i].equalsIgnoreCase("Penalty(€)")){
                    iPenalty = i;                
            }
            else if(header[i].equalsIgnoreCase("Earnback(€)")){
                    iEarn_back = i;
        }
            else {
                    continue;
            }
        }
    XSSFWorkbook w = new XSSFWorkbook();   //above Excel 2003 
    XSSFSheet sheet = w.createSheet("sheet");
    XSSFCellStyle style = w.createCellStyle();
    XSSFDataFormat df = w.createDataFormat();
    style.setDataFormat(df.getFormat("###,###,###,###,###,##0.00 €"));
    
    int comLan=0,cloud=0,comWan=0,edc=0,iom=0,fmoEwpSmart=0,fmoEwpSwiss=0,incCo=0,printing=0,serviceMgmt=0,uhd = 0,total=0;
        for (int i = 0; i < rowNumber; i++) {
            if(datas[i][0].equalsIgnoreCase("Cloud"))
                cloud = i;
            if(datas[i][0].equalsIgnoreCase("EDC"))
                edc = i; 
            if(datas[i][0].equalsIgnoreCase("Service Level Credit Total"))
                total = i;
        }

            
        
    
        for (int i = 0; i < rowNumber; i++) {
        String[] row = datas[i];
        XSSFRow row2 = sheet.createRow(i);            
        for (int j = 0; j < colNumber; j++) {
            try {
                if (i != 0 && (j == 3 || j == 4)){

                        if(i==cloud && (j==iPenalty || j==iEarn_back)){
                        XSSFCell cell = row2.createCell(j);
                        String col = CellReference.convertNumToColString(cell.getColumnIndex());
                        int upper = cloud + 2;
                        int belower = edc;
                        cell.setCellFormula("Sum(" + col + upper + ":" + col + belower + ")"); // a Sum 1-bázisú indext használ, ezért nem elég csak 1-et hozzáadni.
                        XSSFFormulaEvaluator formulaEvaluator = w.getCreationHelper().createFormulaEvaluator();
                        formulaEvaluator.evaluateFormulaCell(cell);
                        row2.getCell(j).setCellStyle(style);
                        continue;
                        }
                
                        if(i==edc && (j==iPenalty || j==iEarn_back)){
                        XSSFCell cell = row2.createCell(j);
                        String col = CellReference.convertNumToColString(cell.getColumnIndex());
                        int upper = edc + 2;
                        int belower = total;
                        cell.setCellFormula("Sum(" + col + upper + ":" + col + belower + ")");
                        XSSFFormulaEvaluator formulaEvaluator = w.getCreationHelper().createFormulaEvaluator();
                        formulaEvaluator.evaluateFormulaCell(cell);
                        row2.getCell(j).setCellStyle(style);
                        continue;
                        }
                        
                        if(i==total && j==iPenalty){
                        XSSFCell cell = row2.createCell(j);
                        String col = CellReference.convertNumToColString(cell.getColumnIndex());
                        comLan++;cloud++;comWan++;edc++;iom++;fmoEwpSmart++;fmoEwpSwiss++;incCo++;printing++;serviceMgmt++;uhd++;
                        cell.setCellFormula("Sum(" + col + cloud + 
                                "," + col + edc + ")");
                        XSSFFormulaEvaluator formulaEvaluator = w.getCreationHelper().createFormulaEvaluator();
                        formulaEvaluator.evaluateFormulaCell(cell);
                        row2.getCell(j).setCellStyle(style);
                        continue;
                        }
                        if(i==total && j==iEarn_back){
                        XSSFCell cell = row2.createCell(j);
                        String col = CellReference.convertNumToColString(cell.getColumnIndex());
                        cell.setCellFormula("Sum(" + col + cloud + 
                                "," + col + edc + ")");
                        XSSFFormulaEvaluator formulaEvaluator = w.getCreationHelper().createFormulaEvaluator();
                        formulaEvaluator.evaluateFormulaCell(cell);
                        row2.getCell(j).setCellStyle(style);
                        continue;
                        }

                        row2.createCell(j).setCellValue(Double.parseDouble(row[j].replace(",","")));
                        row2.getCell(j).setCellStyle(style);

                        continue;
                }else{
                row2.createCell(j).setCellValue(row[j]);
                }
                } catch(NumberFormatException ex) {
                        System.out.println(ex.getMessage());
               } 
                          
                               }
     
  //          row2.createCell(j).setCellValue(row[j].replace("(€)",""));
        }             
    
            for (int i = 0; i < colNumber; i++) {
            w.getSheet("sheet").autoSizeColumn(i);
            
        }
            int [] highlight = {cloud-1,edc-1,total};
            XSSFCellStyle style3 = w.createCellStyle();
            style3.setFillForegroundColor(IndexedColors.GOLD.getIndex());
            style3.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            style3.setDataFormat(df.getFormat("###,###,###,###,###,##0.00 €"));
                for (int i = 0; i < highlight.length; i++) {
                    XSSFRow row3 = w.getSheet("sheet").getRow(highlight[i]);
                    for (int j = 0; j <=iEarn_back; j++) {
                        row3.getCell(j).setCellStyle(style3);
                        
                    }
        }
                try {
                    FileOutputStream output = new FileOutputStream(filePathOut);
                    w.write(output);
                    } catch (IOException ex) {
                    Logger.getLogger(FormatterKlass.class.getName()).log(Level.SEVERE, null, ex);
                }
}
}
