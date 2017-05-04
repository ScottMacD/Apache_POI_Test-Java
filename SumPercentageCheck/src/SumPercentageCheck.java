import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintStream;
import java.math.RoundingMode;
import java.text.DecimalFormat;

import javax.swing.JOptionPane;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import java.util.Random;


public class SumPercentageCheck {
	
	public static void main (String[] args) throws IOException, InterruptedException{
		SumPercentageCheck test = new SumPercentageCheck();
	    test.Scenario1();
        test.Scenario2();
        test.Scenario3();
	}
	
	  public void Scenario1() throws IOException  {
	     FileInputStream ExcelTest1 = new FileInputStream(new File("C:/Users/Family/Desktop/TestData/Product Volumes.xls"));
         HSSFWorkbook wb = new HSSFWorkbook(ExcelTest1);
		 HSSFSheet sheet = wb.getSheetAt(0);
	
        final int liquidVolume = 3200;
		FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
        CellReference cellReference = new CellReference("B1202"); 
		Row row = sheet.getRow(cellReference.getRow());
		Cell cell = row.getCell(cellReference.getCol());
		CellValue cellValue = evaluator.evaluate(cell);
		
		DecimalFormat df = new DecimalFormat("#.##");
		df.setRoundingMode(RoundingMode.CEILING);
        double sumValue = cellValue.getNumberValue();
        double PercentageFilled = sumValue * 100 / liquidVolume;
        
        File file = new File("C:/Users/Family/Desktop/TestData/Test.html");
 		FileOutputStream fos = new FileOutputStream(file);
 		PrintStream ps = new PrintStream(fos);
 		System.setOut(ps);
        System.out.println("The sum of the volume column is " + df.format(sumValue));
        System.out.println("<br />");
        System.out.println("The percentage filled is " + df.format(PercentageFilled)+"%");
        fos.flush();
        fos.close();
        wb.close();
	}

	  public void Scenario2() throws IOException  {
		     FileInputStream ExcelTest1 = new FileInputStream(new File("C:/Users/Family/Desktop/TestData/Product Volumes.xls"));
	         HSSFWorkbook wb = new HSSFWorkbook(ExcelTest1);
			 HSSFSheet sheet2 = wb.getSheetAt(1);
			 
			 final int liquidVolume = 3200;
		  
	      for (Row rows : sheet2) {
	          for (Cell cellRotate : rows) {
	              switch (cellRotate.getCellType()) {
	                  case Cell.CELL_TYPE_NUMERIC:
	                  System.out.println(cellRotate.getNumericCellValue());
	                      break;
	                  case Cell.CELL_TYPE_BLANK:
	    	        	double randomDig = Math.random()*.5;
                        cellRotate.setCellValue(randomDig);
	                    System.out.println(cellRotate.getNumericCellValue());
	                      break;
	                  default:
	                      System.out.println();
	                      
	              }
	          }  
	      }
	      
	      	FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
	         CellReference cellReference = new CellReference("B8335"); 
			 Row row = sheet2.getRow(cellReference.getRow());
			 Cell cell = row.getCell(cellReference.getCol());
			 CellValue cellValue = evaluator.evaluate(cell);
				
			 DecimalFormat df = new DecimalFormat("#.##");
			 df.setRoundingMode(RoundingMode.CEILING);
		     double sumValue = cellValue.getNumberValue();
		     double PercentageFilled = sumValue * 100 / liquidVolume;
	      
	        File file = new File("C:/Users/Family/Desktop/TestData/Test2.html");
			  FileOutputStream fos = new FileOutputStream(file);
			  PrintStream ps = new PrintStream(fos);
			 System.setOut(ps);
	         System.out.println("The sum of the volume column is " + df.format(sumValue));
	         System.out.println("<br />");
	         System.out.println("The percentage filled is " + df.format(PercentageFilled)+"%");
	         fos.flush();
	         fos.close();
	         wb.close();
	        }
	  
	  public void Scenario3() throws IOException  {
		     FileInputStream ExcelTest1 = new FileInputStream(new File("C:/Users/Family/Desktop/TestData/Product Volumes.xls"));
	         HSSFWorkbook wb = new HSSFWorkbook(ExcelTest1);
			  HSSFSheet sheet3 = wb.getSheetAt(2);
	         final int liquidVolume = 3200;
			 FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();

			CellReference cellReference = new CellReference("B2332"); 
			Row row = sheet3.getRow(cellReference.getRow());
			Cell cell = row.getCell(cellReference.getCol());
			CellValue cellValue = evaluator.evaluate(cell);
	        double sumValue = cellValue.getNumberValue();
	        DecimalFormat df = new DecimalFormat("#.##");
			df.setRoundingMode(RoundingMode.CEILING);
	        double overageAmt = sumValue - liquidVolume;
	        
	        if(sumValue >= liquidVolume){
	        	JOptionPane.showMessageDialog(null, "The volume amount was over by " + df.format(overageAmt) + " cubic feet");
		     wb.close();
	        }
	   }
}