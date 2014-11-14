package excelStickerFormater;

import java.awt.EventQueue;

import javax.swing.DefaultListModel;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JTextField;
import javax.swing.JButton;
import javax.swing.JLabel;
import javax.swing.JList;
import javax.swing.SwingConstants;

import java.awt.event.ActionListener;
import java.awt.event.ActionEvent;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import javax.swing.AbstractListModel;
import javax.swing.JScrollPane;

import jxl.Cell;
import jxl.CellType;
import jxl.Sheet;
import jxl.Workbook;
import jxl.format.Alignment;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableCell;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

import java.awt.Color;
import java.awt.Font;

import javax.swing.JComboBox;
import javax.swing.DefaultComboBoxModel;

public class ExcelSticker {
	final static int HEADER_ROW =2;
	final static int RMAX =50;
	final static int CMAX =30;
	private JFrame frame;
	private JTextField txfDataPath;
	private JTextField txfBlankPath;
	private JLabel lblCount;
	private JLabel lblMsg;
	private JLabel lblFormat;
	private Workbook ro_blank_workbook = null;
	private Workbook ro_data_workbook = null;
	private Sheet ro_sheet = null;
	private DefaultListModel listModel = new DefaultListModel();
	private JList<String> lstProject = new JList(listModel);
	File blankfile = null;
	JButton btnLoadData;
	JButton btnLoadTemplate;
	JButton btnGenerate;
	JComboBox cbxColReadCount = new JComboBox();
	List<ArrayList<String>> stList = null; 
	
	private int TEMP_COLS = 5;
	private int TEMP_ROWS = 12;
	private int proj_count = 0;

	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					ExcelSticker window = new ExcelSticker();
					window.frame.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	public ExcelSticker() {
		initialize();
	}


	private void initialize() {
		frame = new JFrame();
		frame.setTitle("\u751F\u7522\u8A08\u5283\u6A19\u7C64\u7522\u751F\u5668");
		frame.setBounds(100, 100, 546, 426);
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frame.getContentPane().setLayout(null);
		final JFileChooser fc = new JFileChooser();
		
		txfDataPath = new JTextField();
		txfDataPath.setFont(new Font("PMingLiU", Font.PLAIN, 14));
		txfDataPath.setBounds(235, 119, 289, 21);
		frame.getContentPane().add(txfDataPath);
		txfDataPath.setColumns(10);
		

		
		JLabel label = new JLabel("\u8A08\u756B\u6E05\u55AE");
		label.setFont(new Font("PMingLiU", Font.PLAIN, 14));
		label.setBounds(10, 170, 70, 15);
		frame.getContentPane().add(label);
		
		JLabel label_1 = new JLabel("\u8A08\u756B\u6578\u91CF:");
		label_1.setFont(new Font("PMingLiU", Font.PLAIN, 14));
		label_1.setBounds(178, 143, 61, 15);
		frame.getContentPane().add(label_1);
		
		lblCount = new JLabel("");
		lblCount.setForeground(Color.MAGENTA);
		lblCount.setFont(new Font("PMingLiU", Font.BOLD, 12));
		lblCount.setHorizontalAlignment(SwingConstants.LEFT);
		lblCount.setBounds(245, 143, 85, 15);
		frame.getContentPane().add(lblCount);
		
		lblMsg = new JLabel("");
		lblMsg.setFont(new Font("PMingLiU", Font.BOLD, 14));
		lblMsg.setForeground(Color.RED);
		lblMsg.setBounds(0, 366, 530, 21);
		frame.getContentPane().add(lblMsg);


		
		
		btnLoadData = new JButton("3. \u8B80\u53D6\u8A08\u756B\u6A94");
		btnLoadData.setFont(new Font("PMingLiU", Font.PLAIN, 14));
		btnLoadData.setEnabled(false);
		btnLoadData.setHorizontalAlignment(SwingConstants.LEFT);
		btnLoadData.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
		        int returnVal = fc.showOpenDialog(frame);

		        if (returnVal == JFileChooser.APPROVE_OPTION) {
		            File file = fc.getSelectedFile();
		            txfDataPath.setText(file.getAbsolutePath());
		            
		            try {
		            	ro_data_workbook = Workbook.getWorkbook(file);
					} catch (BiffException e) {
						lblMsg.setText("檔案格式錯誤!");
					} catch (IOException e) {
						lblMsg.setText("無法存取檔案!");
					}
		            
		            processExcel();
		            
		        } else {
		        }
			}
		});
		btnLoadData.setBounds(10, 120, 158, 40);
		frame.getContentPane().add(btnLoadData);
		
		btnGenerate = new JButton("4. \u7522\u751F\u6A19\u7C64\u6A94");
		btnGenerate.setFont(new Font("PMingLiU", Font.PLAIN, 14));
		btnGenerate.setEnabled(false);
		btnGenerate.setHorizontalAlignment(SwingConstants.LEFT);
		btnGenerate.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				processGenerate();
				
			}
		});
		btnGenerate.setBounds(14, 316, 158, 39);
		frame.getContentPane().add(btnGenerate);
		
		JLabel label_2 = new JLabel("檔案");
		label_2.setFont(new Font("PMingLiU", Font.PLAIN, 14));
		label_2.setBounds(179, 123, 36, 15);
		frame.getContentPane().add(label_2);
		
		JScrollPane scrollPane = new JScrollPane();
		scrollPane.setBounds(10, 195, 510, 111);
		frame.getContentPane().add(scrollPane);
		lstProject.setFont(new Font("PMingLiU", Font.PLAIN, 14));
		
		
		scrollPane.setViewportView(lstProject);
		
		JButton btnExit = new JButton("離開");
		btnExit.setFont(new Font("PMingLiU", Font.PLAIN, 14));
		btnExit.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				System.exit(0);
			}
		});
		btnExit.setBounds(437, 316, 87, 39);
		frame.getContentPane().add(btnExit);
		
		btnLoadTemplate = new JButton("1. 讀取空白標籤Excel");
		btnLoadTemplate.setFont(new Font("PMingLiU", Font.PLAIN, 14));
		btnLoadTemplate.setHorizontalAlignment(SwingConstants.LEFT);
		btnLoadTemplate.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
		        int returnVal = fc.showOpenDialog(frame);

		        if (returnVal == JFileChooser.APPROVE_OPTION) {
		            blankfile = fc.getSelectedFile();
		            txfBlankPath.setText(blankfile.getAbsolutePath());
		            
		            try {
		            	ro_blank_workbook = Workbook.getWorkbook(blankfile);
					} catch (BiffException e) {
						lblMsg.setText("檔案格式錯誤!");
					} catch (IOException e) {
						lblMsg.setText("無法存取檔案!");
					}
	            
		            processBlank();
		            
		        } else {
		        }				
			}
		});
		btnLoadTemplate.setBounds(10, 21, 179, 40);
		frame.getContentPane().add(btnLoadTemplate);
		
		JLabel label_3 = new JLabel("規格:");
		label_3.setFont(new Font("PMingLiU", Font.PLAIN, 14));
		label_3.setBounds(199, 44, 40, 15);
		frame.getContentPane().add(label_3);
		
		lblFormat = new JLabel("");
		lblFormat.setForeground(Color.MAGENTA);
		lblFormat.setFont(new Font("PMingLiU", Font.BOLD, 12));
		lblFormat.setBounds(235, 44, 100, 15);
		frame.getContentPane().add(lblFormat);
		
		JLabel label_4 = new JLabel("\u6A94\u6848");
		label_4.setFont(new Font("PMingLiU", Font.PLAIN, 14));
		label_4.setBounds(199, 25, 40, 15);
		frame.getContentPane().add(label_4);
		
		txfBlankPath = new JTextField();
		txfBlankPath.setFont(new Font("PMingLiU", Font.PLAIN, 14));
		txfBlankPath.setColumns(10);
		txfBlankPath.setBounds(235, 23, 289, 21);
		frame.getContentPane().add(txfBlankPath);
		

		cbxColReadCount.setFont(new Font("PMingLiU", Font.PLAIN, 14));
		cbxColReadCount.setModel(new DefaultComboBoxModel(new String[] {"5", "4", "3", "2", "1"}));
		cbxColReadCount.setSelectedIndex(1);
		cbxColReadCount.setBounds(186, 88, 80, 21);
		frame.getContentPane().add(cbxColReadCount);
		
		JLabel label_5 = new JLabel("2. \u8A2D\u5B9A\u8B80\u53D6\u6B04\u4F4D\u6578");
		label_5.setFont(new Font("PMingLiU", Font.PLAIN, 14));
		label_5.setBounds(16, 91, 156, 15);
		frame.getContentPane().add(label_5);
		
		
	}
	
	private void processBlank(){

        if (ro_blank_workbook != null){
        	ro_sheet = ro_blank_workbook.getSheet(0);
        	
        	int rows = ro_sheet.getRows();
        	int cols = ro_sheet.getColumns();
        	if (rows > 0 && rows < RMAX && cols > 0 && cols < CMAX){
	        	TEMP_ROWS = rows;
	        	TEMP_COLS = cols;
	        	lblFormat.setText(TEMP_COLS + " X " + TEMP_ROWS);
	        	btnLoadData.setEnabled(true);
	        	lblMsg.setText("");
        	}
        	else{
        		lblMsg.setText("超出範圍");
        		lblFormat.setText("");
        	}
        	
        	String test = ro_sheet.getCell(0,0).getContents(); 
        	if (test.length()>0){
        		lblMsg.setText("檔案只能有框線，不能有任何內容!");
        		lblFormat.setText("");
        	}
        	
        }
        else{
        	btnLoadData.setEnabled(false);
        }
	}
	
	private void processExcel(){
        if (ro_data_workbook != null){
        	listModel.clear();
        	//listModel = new DefaultListModel();
        	//lstProject = new JList(listModel);
        	
        	ArrayList<String> entry = new ArrayList<String>(); //each entry has multiple columns
        	
        	stList = new ArrayList<ArrayList<String>>();
        	
        	Cell cell = null;
        	
        	ro_sheet = ro_data_workbook.getSheet(0);
        	
        	int rows = ro_sheet.getRows();
        	int cols = ro_sheet.getColumns();
        	
        	proj_count=0;
        	for (int r=2;r<rows;r++){
        		String dataString = "";
        		entry  = new ArrayList<String>();
        		
        		//read rows
        		Cell crow[] = ro_sheet.getRow(r);
        		
        		//read columns 
        		for (int cidx=0; cidx < Integer.valueOf((String)cbxColReadCount.getSelectedItem()); cidx++){
        			entry.add(crow[cidx].getContents());
        			dataString = dataString + crow[cidx].getContents() + " ; ";
        		}
        		        		
        		if (entry.size() > 0){
        			proj_count++;
        			listModel.addElement(proj_count + ": " + dataString);
        			stList.add(entry);
        		}
        		else{
        			lblMsg.setText("資料檔讀取異常, 請聯絡#2716");
        		}
        	}
        	
        	lblCount.setText(""+listModel.size());
        	
        	if (proj_count>0){
        		btnGenerate.setEnabled(true);
        		lblMsg.setText("");
        	}
            else{
            	btnGenerate.setEnabled(false);
            }        	
        }
		
	
	}
	
	public class IndexClass{
		public int page =0;
		public int row=0;
		public int col=0;
		
		public IndexClass(){}
	}
	
	private IndexClass getIndex(int projnum){
		int proj_page = TEMP_COLS * TEMP_ROWS;
		int page = 0;
		int col=0,row=0,tmpcol=0,tmprow=0;
		IndexClass idx = new IndexClass();
		
		page = projnum / proj_page; //取整數
		if (projnum % proj_page > 0)
			page++;
	
		tmprow = projnum - (proj_page * (page-1));
		row = tmprow / TEMP_COLS;
		if (tmprow % TEMP_COLS == 0)
			row--;
		
		tmpcol = tmprow % TEMP_COLS;
		if (tmpcol == 0) 
			col = TEMP_COLS-1;
		else
			col = tmpcol-1;

		idx.page = page;
		idx.row = row;
		idx.col = col;
		
		return idx;
	}
	
	void processGenerate(){
		IndexClass idx = null;
		WritableCell cell = null;
		WritableWorkbook copyworkbook = null;
		WritableSheet sheet = null;
		
		List<File> fileList = new ArrayList<File>();
		List<Workbook> ro_workbookList = new ArrayList<Workbook>();
		List<WritableWorkbook> workbookList = new ArrayList<WritableWorkbook>();
		List<WritableSheet> sheetList =  new ArrayList<WritableSheet>();
				
		String filename = "";
		int proj_page = TEMP_COLS * TEMP_ROWS;
		//proj_count = 200; //debug
		int fcount = proj_count/proj_page; //fcount: 檔案個數
		if (proj_count % proj_page != 0){ //有餘數
			fcount++;
		}
			
		lblMsg.setText("將產生"+fcount+"個標籤印刷檔.");

		for (int f=0; f<fcount; f++){
			try {
				
				ro_workbookList.add (  Workbook.getWorkbook( blankfile ));
			} catch (BiffException e1) {	e1.printStackTrace();} 
			catch (IOException e1) {	e1.printStackTrace();}
			
			
			
			filename = "output"+ (f+1) + ".xls";
			
			try {
				workbookList.add( Workbook.createWorkbook(new File(filename), ro_workbookList.get(f)) );
			} catch (IOException e) {e.printStackTrace();	}
			
			sheetList.add(workbookList.get(f).getSheet(0));
		}
		
		System.out.println("sheetCount:"+ sheetList.size());
		
		
		ArrayList<String> entry1 = stList.get(1);
		System.out.println("rec1:" + entry1.get(0));
		ArrayList<String> entry2 = stList.get(2);
		System.out.println("rec2:" + entry2.get(0));
		
		String brk = "\r\n";
		for (int pc=1; pc<=proj_count; pc++){
			ArrayList<String> entry = stList.get(pc-1); //each entry has multiple columns
			String data = "";
			
			for (int cidx=0; cidx < entry.size(); cidx++){
				if (cidx < entry.size()-1){
					data = data + entry.get(cidx) + brk;
				}else{
					data = data + entry.get(cidx);
				}
			}

			
			idx = getIndex(pc);
			
			System.out.println(pc+".  p:" + idx.page + "  r:" + idx.row + "  c:" +idx.col);
			
			WritableFont arial10font = new WritableFont(WritableFont.ARIAL, 10); 
			WritableCellFormat arial10format = new WritableCellFormat (arial10font); 
			
			try {
				arial10format.setAlignment(Alignment.CENTRE);
			} catch (WriteException e2) {	e2.printStackTrace();	}
			
			try {
				arial10format.setWrap(true);
			} catch (WriteException e1) {e1.printStackTrace();	}

			
			WritableCell  c = sheetList.get(idx.page-1).getWritableCell(idx.row, idx.col);
			//if (c.getType()== CellType.EMPTY){
				//System.out.println("it's empty");
				Label l = new Label(idx.col,idx.row, data, arial10format);
				try {
					sheetList.get(idx.page-1).addCell(l);
				} catch (RowsExceededException e) {	e.printStackTrace();} 
				  catch (WriteException e) {e.printStackTrace();}
				
			/*
			  }else if(c.getType()== CellType.LABEL){
				System.out.println("it's label");
				Label l = (Label) c;
				l.setString(data);
			}
			*/
			
		}
		
		for (int f=0; f<fcount; f++){
			
			try {
				workbookList.get(f).write();
				System.out.println("write file"+f);
			} catch (IOException e) {e.printStackTrace();}
			
			try {
				workbookList.get(f).close();
			} catch (WriteException e) {e.printStackTrace();
			} catch (IOException e) {e.printStackTrace();}
		}
		
		
	}
}
