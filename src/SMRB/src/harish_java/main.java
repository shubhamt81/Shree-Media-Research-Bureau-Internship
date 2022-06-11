package harish_java;
//
////import javax.swing.*;
////import java.awt.*;
//import java.awt.FlowLayout;
//import java.awt.event.ActionEvent;
//import java.awt.event.ActionListener;
//
//import javax.swing.JButton;  
//import javax.swing.JFrame;  
//import javax.swing.JLabel;  
////import javax.swing.JPanel;
////import javax.swing.JTextField;
//import javax.swing.JComboBox;
//
//  
//class  CasteData implements ActionListener {  
//	JFrame frame = new JFrame("Student data INPUT");  
////    JPanel panel = new JPanel();
//    JLabel l1=new JLabel("SIRNAME");
////    JTextField t1=new JTextField(10);
//    JLabel l2=new JLabel("CASTE");
//    JComboBox c1=new JComboBox();
//    JComboBox c2=new JComboBox();
//    JButton b1 = new JButton();  
//    JButton b2 = new JButton(); 
//    
//    
//    
//    CasteData(){  
////    	
//	    b1.addActionListener((ActionListener) this);
//	    frame.setSize(400, 300);
//	    b1.setText("SET");
//	    b2.setText("RESET");
//	    c2.insertItemAt("Bhramin",0);
//	    c2.insertItemAt("Maratha",1);
//	    c2.insertItemAt("Sonar",2);
//	    c2.insertItemAt("Parssi",3);
//	    c2.insertItemAt("Banjara",4);
//	    frame.add(l1);
//	    frame.add(c1);
//	    frame.add(l2);
//	    frame.add(c2);
//	    frame.add(b1);
//	    frame.add(b2);
//	    
//	    
//	    frame.setLayout(new FlowLayout(FlowLayout.LEFT,100,40));
//	    frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);  
//	    frame.setVisible(true); 
//
//    }
//    public void actionPerformed(ActionEvent e) {
//    	JFrame f1=new JFrame();
//    	f1.setSize(450,300);
//    	f1.add(new JLabel("WORKING"));
//    	f1.setVisible(true);
//    }
//    
//} 
//public class main {
//	public static void main(String args[]) {
//		new CasteData();
//	}
//}


import java.io.File;
import java.io.FileInputStream;
//import java.io.IOException;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.*;
// Main class
public class GFG {
 
    // Main driver method
    public static void main(String[] args)
    {
 
        // Try block to check fo exceptions
        try {
 
            // Reading file fro local directory
            FileInputStream file = new FileInputStream(
                new File("gfgcontribute.xlsx"));
 
            // Create Workbook instance holding reference to
            // .xlsx file
            XSSFWorkbook workbook = new XSSFWorkbook(file);
 
            // Get first/desired sheet from the workbook
            XSSFSheet sheet = workbook.getSheetAt(0);
 
            // Iterate through each rows one by one
            Iterator<Row> rowIterator = sheet.iterator();
 
            // Till there is an element condition holds true
            while (rowIterator.hasNext()) {
 
                Row row = rowIterator.next();
 
                // For each row, iterate through all the
                // columns
                Iterator<Cell> cellIterator
                    = row.cellIterator();
 
                while (cellIterator.hasNext()) {
 
                    Cell cell = cellIterator.next();
 
                    // Checking the cell type and format
                    // accordingly
                    switch (cell.getCellType()) {
 
                    // Case 1
                    case Cell.CELL_TYPE_NUMERIC:
                        System.out.print(
                            cell.getNumericCellValue()
                            + "t");
                        break;
 
                    // Case 2
                    case Cell.CELL_TYPE_STRING:
                        System.out.print(
                            cell.getStringCellValue()
                            + "t");
                        break;
                    }
                }
 
                System.out.println("");
            }
 
            // Closing file output streams
            file.close();
        }
 
        // Catch block to handle exceptions
        catch (Exception e) {
 
            // Display the exception along with line number
            // using printStackTrace() method
            e.printStackTrace();
        }
    }
}