package breakdowner;

import java.io.*;
import java.awt.*;
import java.awt.event.*;
import javax.swing.*;
import javax.swing.SwingUtilities;
import javax.swing.filechooser.*;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class FileChooser extends JPanel
                             implements ActionListener {
    static private final String newline = "\n";
    JButton openButton, saveButton, dirButton;
    JTextArea log;
    JFileChooser fc;
    JFileChooser dc;
    File[] files;
    File dir;
    String filename;
    //String dirName;
    
    static int ARCH_TOTAL = 259;
    static int CON_TOTAL = 127;
    static int VM_TOTAL = 115;
    static int NS_TOTAL = 70;
    
    
    
    
    
    
    static int A_Delivery;
    static int A_Type;
    static int A_SKU;
    static int A_Description;
    static int A_QTY;
    
    static int C_Delivery;
    static int C_Type;
    static int C_SKU;
    static int C_Description;
    static int C_QTY;
    
    static int V_Delivery;
    static int V_Type;
    static int V_SKU;
    static int V_Description;
    static int V_QTY;
    
    static int N_Delivery;
    static int N_Type;
    static int N_SKU;
    static int N_Description;
    static int N_QTY;
    

    public FileChooser() {
        super(new BorderLayout());

        //Create the log first, because the action listeners
        //need to refer to it.
        log = new JTextArea(5,20);
        log.setMargin(new Insets(5,5,5,5));
        log.setEditable(false);
        JScrollPane logScrollPane = new JScrollPane(log);

        //Create a file chooser
        fc = new JFileChooser();
        fc.setMultiSelectionEnabled(true);

        dc = new JFileChooser();
        dc.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        //fc.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES);

        
        openButton = new JButton("Open file");
        openButton.addActionListener(this);

      
        saveButton = new JButton("Create breakdown(s)");
        saveButton.addActionListener(this);
        
        dirButton = new JButton("Choose output directory");
        dirButton.addActionListener(this);

        //For layout purposes, put the buttons in a separate panel
        JPanel buttonPanel = new JPanel(); //use FlowLayout
        buttonPanel.add(openButton);
        buttonPanel.add(saveButton);
        buttonPanel.add(dirButton);

        dir = new File("");
        
        //Add the buttons and the log to this panel.
        add(buttonPanel, BorderLayout.PAGE_START);
        add(logScrollPane, BorderLayout.CENTER);
    }

    public void actionPerformed(ActionEvent e) {

        //Handle open button action.
    	if(e.getSource() == dirButton){
    		int returnVal = dc.showOpenDialog(FileChooser.this);
    		
    		if(returnVal == JFileChooser.APPROVE_OPTION) {
    			dir = dc.getSelectedFile();
    			log.append("Output directory: " + dir.toString());
    			
    			
    			
    		} else {
    			log.append("Directory command aborted." + newline);
    		}
    		
    		
    	}
    	
        if (e.getSource() == openButton) {
            int returnVal = fc.showOpenDialog(FileChooser.this);

            if (returnVal == JFileChooser.APPROVE_OPTION) {
                files = fc.getSelectedFiles();
                
                
                //List files selected
                for(File file : files)
                {
                	              
                filename = file.getName();
                log.append("Opening: " + file.getName() + ", ");
                
                }
                log.append(newline);
            } else {
                log.append("Open command aborted." + newline);
            }
            log.setCaretPosition(log.getDocument().getLength());

        
            
            
        
        //Handle save button action.
        } else if (e.getSource() == saveButton) {
        	files = fc.getSelectedFiles();
            for(File file : files)
            {
                //file = fc.getSelectedFile();
                //This is where the magic happens *_*
                log.append("Converting: " + file.getName() + "." + newline);
                
                String filename = file.getName();
               
                
                Workbook inputBook = null;
        		try {
        			inputBook = WorkbookFactory.create(file);
        		} catch (EncryptedDocumentException e1) {
        			// TODO Auto-generated catch block
        			e1.printStackTrace();
        		} catch (InvalidFormatException e1) {
        			// TODO Auto-generated catch block
        			e1.printStackTrace();
        		} catch (IOException e1) {
        			// TODO Auto-generated catch block
        			e1.printStackTrace();
        		}
        		
        		//FIND THE COLUMNS BRO
        		Sheet inputSheet = inputBook.getSheetAt(1);		//grab arch sheet first
        	    Row inRow = inputSheet.getRow(6);	// row 7 headers
        	    for(int i=0; i<inRow.getLastCellNum(); i++)
        	    {
        	    	String cellValue = inRow.getCell(i).getStringCellValue();
        	    	if( cellValue.equalsIgnoreCase("Delivery") )
        	    		A_Delivery = i;
        	    	if( cellValue.equalsIgnoreCase("Fixture Type"))
        	    		A_Type = i;
        	    	if( cellValue.equalsIgnoreCase("SKU") )
        	    		A_SKU = i;
        	    	if( cellValue.equalsIgnoreCase("Fixture Item"))
        	    		A_Description = i;
        	    	if( cellValue.equalsIgnoreCase("QTY") )
            	    	A_QTY = i;
            	    	
        	    }
        	    
        	    inputSheet = inputBook.getSheetAt(2);		//construction sheet
        	    inRow = inputSheet.getRow(6);
        	    for(int i=0; i<inRow.getLastCellNum(); i++)
        	    {
        	    	String cellValue = inRow.getCell(i).getStringCellValue();
        	    	if( cellValue.equalsIgnoreCase("Delivery #") )
        	    		C_Delivery = i;
        	    	if( cellValue.equalsIgnoreCase("Material Type"))
        	    		C_Type = i;
        	    	if( cellValue.equalsIgnoreCase("SKU") )
        	    		C_SKU = i;
        	    	if( cellValue.equalsIgnoreCase("Fixture Item"))
        	    		C_Description = i;
        	    	if( cellValue.equalsIgnoreCase("QTY") )
            	    	C_QTY = i;
            	    	
        	    }
        	    
        	    inputSheet = inputBook.getSheetAt(3);		//VM sheet
        	    inRow = inputSheet.getRow(6);
        	    for(int i=0; i<inRow.getLastCellNum(); i++)
        	    {
        	    	String cellValue = inRow.getCell(i).getStringCellValue();
        	    	if( cellValue.equalsIgnoreCase("Delivery #") )
        	    		V_Delivery = i;
        	    	if( cellValue.equalsIgnoreCase("Fixture Type"))
        	    		V_Type = i;
        	    	if( cellValue.equalsIgnoreCase("SKU") )
        	    		V_SKU = i;
        	    	if( cellValue.equalsIgnoreCase("Fixture Item"))
        	    		V_Description = i;
        	    	if( cellValue.equalsIgnoreCase("QTY") )
            	    	V_QTY = i;
        	    }
        	    
        	    inputSheet = inputBook.getSheetAt(4);		//non standard sheet
        	    inRow = inputSheet.getRow(6);
        	    for(int i=0; i<inRow.getLastCellNum(); i++)
        	    {
        	    	String cellValue = inRow.getCell(i).getStringCellValue();
        	    	if( cellValue.equalsIgnoreCase("Delivery #") )
        	    		N_Delivery = i;
        	    	if( cellValue.equalsIgnoreCase("Fixture Type"))
        	    		N_Type = i;
        	    	if( cellValue.equalsIgnoreCase("SKU") )
        	    		N_SKU = i;
        	    	if( cellValue.equalsIgnoreCase("Fixture Item"))
        	    		N_Description = i;
        	    	if( cellValue.equalsIgnoreCase("QTY") )
            	    	N_QTY = i;
        	    }
        		
        		System.out.println("A_Delivery: "+ A_Delivery);
        		System.out.println("A_Type: "+ A_Type);
        		System.out.println("A_SKU: "+ A_SKU);
        		System.out.println("A_Description: "+ A_Description);
        		System.out.println("A_QTY: " + A_QTY);
        		
        		
        		System.out.println("C_Delivery: "+ C_Delivery);
        		System.out.println("C_Type: "+ C_Type);
        		System.out.println("C_SKU: "+ C_SKU);
        		System.out.println("C_Description: "+ C_Description);
        		System.out.println("C_QTY: " + C_QTY);
        		
        		
        		System.out.println("V_Delivery: "+ V_Delivery);
        		System.out.println("V_Type: "+ V_Type);
        		System.out.println("V_SKU: "+ V_SKU);
        		System.out.println("V_Description: "+ V_Description);
        		System.out.println("V_QTY: " + V_QTY);
        		
        		
        		
        		
        		
        		
        		
        		
        		FormulaEvaluator evaluator = inputBook.getCreationHelper().createFormulaEvaluator();
        		
        		
        		Workbook wb = new XSSFWorkbook();		//create new workbook
        		XSSFSheet outSheet = (XSSFSheet) wb.createSheet();		//create new sheet in wb
        		
        		Row row = outSheet.createRow(0);		//create initial row
        	    Cell cell = row.createCell(0);			//create cell at 0x0 (1,1 in excel)
        	    
        	    // START CREATE TITLE ROW
        	    
        	    inputSheet = inputBook.getSheetAt(0);	//grab first sheet to get store name
        	    
        	    /*
        	     * Version 1.1 totals
        	     * 		arch	274
        	     * 		con		125
        	     * 		vm		113
        	     * 		ns		65
        	     * 
        	     * 		row 52 col 1 (51,0) Version: 1.1 - 6/16/15
        	     * 
        	     * 
        	     * NOT VERSION 1.1
        	     * 		arch 	259
        	     * 		con		127
        	     * 		vm		115
        	     * 		NS		70
        	     */
        	    
        	   
        	    
        	    ARCH_TOTAL=259;
        	    CON_TOTAL = 127;
        	    VM_TOTAL = 115;
        	    NS_TOTAL = 65;
        	    
        	    System.out.println(inputSheet.getLastRowNum());
        	    //System.out.println("ROW 51 last cell num:" + inRow.getLastCellNum());
        	    if(inputSheet.getLastRowNum() >=51){
        	    	 inRow = inputSheet.getRow(51);
        	    	if(inRow.getCell(0).getStringCellValue().equalsIgnoreCase("Version: 1.1 - 6/16/15"))
        	    	{
        	    		ARCH_TOTAL = 274;
        	    		CON_TOTAL = 125;
        	    		VM_TOTAL = 113;
        	    		NS_TOTAL = 69;
        	    	}
        	    }
        	    
        	    
        	    inRow = inputSheet.getRow(15);
        	    
        	    String storeName = inRow.getCell(1).toString();
        	    
        	    
        	    Font font = wb.createFont();
        	    font.setFontHeightInPoints((short)20);
        	    font.setFontName("Calibri");
        	    font.setBold(true);
        	    font.setColor(IndexedColors.WHITE.getIndex());
        	    
        	    CellStyle titleStyle = wb.createCellStyle();
        	    titleStyle.setFont(font);
        	    titleStyle.setFillForegroundColor(IndexedColors.DARK_BLUE.getIndex());
        	    titleStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
        	    titleStyle.setAlignment(CellStyle.ALIGN_CENTER);
        	    
        	    
        	    cell.setCellValue(storeName);
        	    
        	    
        	    cell.setCellStyle(titleStyle);
        	    // END SET TITLE ROW
        	    
        	    
        	    // START CREATE HEADER ROW
        	    row = outSheet.createRow(2);		//create row 3
        	    row.createCell(0).setCellValue("Delivery");
        	    row.createCell(1).setCellValue("Type");
        	    row.createCell(2).setCellValue("SKU");
        	    row.createCell(3).setCellValue("Item Description");
        	    row.createCell(4).setCellValue("Source");
        	    row.createCell(5).setCellValue("QTY");
        	    row.createCell(6).setCellValue("Status");
        	    row.createCell(7).setCellValue("Missing Item");
        	    row.createCell(8).setCellValue("Missing QTY");
        	    row.createCell(9).setCellValue("Missing Item Description");
        	    row.createCell(10).setCellValue("O,S,D");
        	    row.createCell(11).setCellValue("IFR Report");
        	    row.createCell(12).setCellValue("Vizona Comment");
        	    
        	    Font font2 = wb.createFont();
        	    font2.setFontHeightInPoints((short)12);
        	    font2.setFontName("Calibri");
        	    font2.setBold(true);
        	    
        	    
        	    CellStyle headerStyle = wb.createCellStyle();
        	    headerStyle.setFont(font);
        	    
        	    row.setRowStyle(headerStyle);
        	    
        	    // END CREATE HEADER ROW
        	    

        	    outSheet.addMergedRegion(new CellRangeAddress(
        	            0, //first row (0-based)
        	            1, //last row  (0-based)
        	            0, //first column (0-based)
        	            12  //last column  (0-based)
        	    ));
        		
        		int rowCounter =3;//start at row 4
        		
        		
        		
        		inputSheet = inputBook.getSheetAt(1);
        						
        		for(int i=7; i<ARCH_TOTAL; i++)						//ARCHITECTURE loop
        		{
        		
        			
        		Row inputRow = inputSheet.getRow(i);	//ROW NUMBER-1
        		Row outRow = outSheet.createRow(rowCounter);
        		if(inputRow.getCell(1).toString().equalsIgnoreCase("") == false)
        		{
        		
        		
        		Cell tempCell = inputRow.getCell(A_Delivery);	//should be delivery #
        		outRow.createCell(0).setCellValue((int)Double.parseDouble(tempCell.toString()));
        		
        		tempCell = inputRow.getCell(A_Type);
        		outRow.createCell(1).setCellValue(tempCell.toString()); //type
        		
        		tempCell = inputRow.getCell(A_SKU);
        		outRow.createCell(2).setCellValue(tempCell.toString());	//sku
        		
        		tempCell = inputRow.getCell(A_Description);
        		outRow.createCell(3).setCellValue(tempCell.toString());	//description
        		
        		tempCell = inputRow.getCell(A_QTY);
        		if(tempCell.getCellType() == Cell.CELL_TYPE_FORMULA)
        		{
        		CellValue cellValue = evaluator.evaluate(tempCell);
        		outRow.createCell(5).setCellValue((int)cellValue.getNumberValue());	//quantity
        		}else{
        			outRow.createCell(5).setCellValue((int) Double.parseDouble(tempCell.toString()));
        		}
        		rowCounter++;
        		}//end if
        		
        		
        		}
        		
        		inputSheet = inputBook.getSheetAt(2);			//SWITCH SHEET
        		
        		for(int i=8; i<CON_TOTAL; i++)						//construction loop
        		{
        		
        			
        		Row inputRow = inputSheet.getRow(i);	//ROW NUMBER-1
        		Row outRow = outSheet.createRow(rowCounter);
        		if(inputRow.getCell(1).toString().equalsIgnoreCase("") == false) //check empty by delivery #
        		{
        		
        		
        		Cell tempCell = inputRow.getCell(C_Delivery);	//should be delivery #
        		outRow.createCell(0).setCellValue((int)Double.parseDouble(tempCell.toString()));
        		
        		tempCell = inputRow.getCell(C_Type);
        		outRow.createCell(1).setCellValue(tempCell.toString()); //type
        		
        		tempCell = inputRow.getCell(C_SKU);
        		outRow.createCell(2).setCellValue(tempCell.toString());	//sku
        		
        		tempCell = inputRow.getCell(C_Description);
        		outRow.createCell(3).setCellValue(tempCell.toString());	//description
        		//System.out.println(tempCell.toString());
        		
        		tempCell = inputRow.getCell(C_QTY);
        		
        		if(tempCell.getCellType() == Cell.CELL_TYPE_FORMULA)
        		{
        		CellValue cellValue = evaluator.evaluate(tempCell);
        		outRow.createCell(5).setCellValue((int)cellValue.getNumberValue());	//quantity
        		}else{
        			outRow.createCell(5).setCellValue(tempCell.toString());
        		}
        		
        		
        		
        		rowCounter++;
        		}//end if
        		
        		
        		}
        		
        		
        		inputSheet = inputBook.getSheetAt(3);	
        		
        		for(int i=8; i<VM_TOTAL; i++)						//VM loop
        		{
        		
        			
        		Row inputRow = inputSheet.getRow(i);	//ROW NUMBER-1
        		Row outRow = outSheet.createRow(rowCounter);
        		if(inputRow.getCell(1).toString().equalsIgnoreCase("") == false) //check empty by delivery #
        		{
        		
        		
        		Cell tempCell = inputRow.getCell(V_Delivery);	//should be delivery #
        		outRow.createCell(0).setCellValue((int)Double.parseDouble(tempCell.toString()));
        		
        		
        		//tempCell = inputRow.getCell(2); unneeded for VM
        		outRow.createCell(1).setCellValue("VM"); //type
        		
        		
        		tempCell = inputRow.getCell(V_SKU);
        		outRow.createCell(2).setCellValue(tempCell.toString());	//sku
        		
        		
        		tempCell = inputRow.getCell(V_Description);
        		outRow.createCell(3).setCellValue(tempCell.toString());	//description
        		
        		
        		tempCell = inputRow.getCell(V_QTY);
        		if(tempCell.getCellType() == Cell.CELL_TYPE_FORMULA)
        		{
        		CellValue cellValue = evaluator.evaluate(tempCell);
        		outRow.createCell(5).setCellValue((int)cellValue.getNumberValue());	//quantity
        		}else{
        			outRow.createCell(5).setCellValue((int) Double.parseDouble(tempCell.toString()));
        		}
        		
        		rowCounter++;
        		}//end if
        		
        		
        		}//end for
        		
        		inputSheet = inputBook.getSheetAt(4);			//Non Standard fixtures
        		
        		for(int i=10; i<NS_TOTAL; i++)						//NON standard loop
        		{
        		
        			
        		Row inputRow = inputSheet.getRow(i);	//ROW NUMBER-1
        		Row outRow = outSheet.createRow(rowCounter);
        		if(inputRow.getCell(1).toString().equalsIgnoreCase("") == false) //check empty by delivery #
        		{
        		
        		
        		Cell tempCell = inputRow.getCell(N_Delivery);	//should be delivery #
        		outRow.createCell(0).setCellValue((int)Double.parseDouble(tempCell.toString()));
        		
        		
        		tempCell = inputRow.getCell(N_Type);
        		outRow.createCell(1).setCellValue(tempCell.toString()); //type
        		
        		
        		tempCell = inputRow.getCell(N_SKU);
        		outRow.createCell(2).setCellValue(tempCell.toString());	//sku
        		
        		
        		tempCell = inputRow.getCell(N_Description);
        		outRow.createCell(3).setCellValue(tempCell.toString());	//description
        		
        		
        		tempCell = inputRow.getCell(N_QTY);
        		if(tempCell.getCellType() == Cell.CELL_TYPE_FORMULA)
        		{
        		CellValue cellValue = evaluator.evaluate(tempCell);
        		outRow.createCell(5).setCellValue((int)cellValue.getNumberValue());	//quantity
        		}else{
        			outRow.createCell(5).setCellValue((int) Double.parseDouble(tempCell.toString()));
        		}
        		
        		rowCounter++;
        		}//end if
        		
        		
        		}//end for
        		
        		
        		
        		
        		
        		
        		
        		for(int i=0; i<13; i++)
        			outSheet.autoSizeColumn(i);
        		
        		/* START TABLE STYLE 
        		XSSFTable myTable = outSheet.createTable();
        		CTTable cttable = myTable.getCTTable();
        		
        						
        		 // Define Styles 
        		   CTTableStyleInfo table_style = cttable.addNewTableStyleInfo();
        		   table_style.setName("TableStyleMedium9"); 
        		   
        		   // Define Style Options 
        		   table_style.setShowColumnStripes(false); //showColumnStripes=0
        		   table_style.setShowRowStripes(true); //showRowStripes=1
           
        		   // Define the data range including headers 
        		   
        		   int lastRow = outSheet.getLastRowNum();
        		   int lastCell = outSheet.getRow(lastRow).getLastCellNum();
        		   
        		   AreaReference my_data_range = new AreaReference(new CellReference(0, 0), new CellReference(lastRow, lastCell));  
        		   
        		   // Set Range to the Table 
        		   cttable.setRef(my_data_range.formatAsString());
        		   cttable.setDisplayName("MYTABLE");      // this is the display name of the table 
        		   cttable.setName("Test");    // This maps to "displayName" attribute in &lt;table&gt;, OOXML             
        		   cttable.setId(1L); //id attribute against table as long value
        		   
        		   
        		   
        		   //Add header columns                
        		   CTTableColumns columns = cttable.addNewTableColumns();
        		   columns.setCount(13L); //define number of columns
        		    //Define Header Information for the Table 
        		    for (int i = 0; i < 13; i++)
        		    {
        		    CTTableColumn column = columns.addNewTableColumn();   
        		    column.setName("Column" + i);      
        		        column.setId(i+1);
        		    }   
        		    
        		
        		
        		//END TABLE STYLE */
        		
        		try {
        			String filePath = dir.getPath();
        			filename = filename.substring(0, filename.length()-5);
        			System.out.println(filePath +"\\" + filename);
        			FileOutputStream fos;
        			if(filePath == "")
        			{
        				fos = new FileOutputStream(filename + " breakdown.xlsx");
        			}else{
        			fos = new FileOutputStream(filePath +"\\"+ filename +" breakdown.xlsx");
        			}
        			wb.write(fos);
        			fos.close();
        			
        			
        		} catch (IOException e1) { e1.printStackTrace(); } 
        		
        		try {
        			wb.close();
        		} catch (IOException e1) {
        			// TODO Auto-generated catch block
        			e1.printStackTrace();
        		}

        	
                
                
                log.append(filename + " breakdown created" + newline);
            }//end for
            } else {
                //log.append("Command cancelled by user." + newline);
            }
            log.setCaretPosition(log.getDocument().getLength());
        
        
    }

    /** Returns an ImageIcon, or null if the path was invalid. */
    protected static ImageIcon createImageIcon(String path) {
        java.net.URL imgURL = FileChooser.class.getResource(path);
        if (imgURL != null) {
            return new ImageIcon(imgURL);
        } else {
            System.err.println("Couldn't find file: " + path);
            return null;
        }
    }

    /**
     * Create the GUI and show it.  For thread safety,
     * this method should be invoked from the
     * event dispatch thread.
     */
    private static void createAndShowGUI() {
        //Create and set up the window.
        JFrame frame = new JFrame("Excelerator 3000 v1.1");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

        //Add content to the window.
        frame.add(new FileChooser());

        //Display the window.
        frame.pack();
        frame.setSize(500,200);
        frame.setLocation(500, 500);
        frame.setVisible(true);
    }
    
    public static String evaluateCell(Cell cell)
    {
    	String answer="";
    	
    	return answer;
    }

    public static void main(String[] args) {
        //Schedule a job for the event dispatch thread:
        //creating and showing this application's GUI.
        SwingUtilities.invokeLater(new Runnable() {
            public void run() {
                //Turn off metal's use of bold fonts
                UIManager.put("swing.boldMetal", Boolean.FALSE); 
                createAndShowGUI();
            }
        });
    }
}
