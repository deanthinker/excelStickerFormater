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
	
	List<StickerClass> stList = null; 
	
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
		frame.setBounds(100, 100, 546, 404);
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frame.getContentPane().setLayout(null);
		final JFileChooser fc = new JFileChooser();
		
		txfDataPath = new JTextField();
		txfDataPath.setBounds(210, 70, 314, 21);
		frame.getContentPane().add(txfDataPath);
		txfDataPath.setColumns(10);
		

		
		JLabel label = new JLabel("\u8A08\u756B\u6E05\u55AE");
		label.setBounds(10, 143, 70, 15);
		frame.getContentPane().add(label);
		
		JLabel label_1 = new JLabel("\u8A08\u756B\u6578\u91CF:");
		label_1.setBounds(178, 94, 61, 15);
		frame.getContentPane().add(label_1);
		
		lblCount = new JLabel("");
		lblCount.setForeground(Color.MAGENTA);
		lblCount.setFont(new Font("PMingLiU", Font.BOLD, 12));
		lblCount.setHorizontalAlignment(SwingConstants.LEFT);
		lblCount.setBounds(240, 93, 55, 15);
		frame.getContentPane().add(lblCount);
		
		lblMsg = new JLabel("");
		lblMsg.setFont(new Font("PMingLiU", Font.BOLD, 14));
		lblMsg.setForeground(Color.RED);
		lblMsg.setBounds(0, 344, 530, 21);
		frame.getContentPane().add(lblMsg);


		
		
		btnLoadData = new JButton("2. 讀取計畫檔");
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
		btnLoadData.setBounds(10, 71, 158, 40);
		frame.getContentPane().add(btnLoadData);
		
		btnGenerate = new JButton("3. 產生標籤檔");
		btnGenerate.setEnabled(false);
		btnGenerate.setHorizontalAlignment(SwingConstants.LEFT);
		btnGenerate.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				processGenerate();
				
			}
		});
		btnGenerate.setBounds(10, 289, 158, 39);
		frame.getContentPane().add(btnGenerate);
		
		JLabel label_2 = new JLabel("檔案");
		label_2.setBounds(179, 74, 36, 15);
		frame.getContentPane().add(label_2);
		
		JScrollPane scrollPane = new JScrollPane();
		scrollPane.setBounds(10, 168, 510, 111);
		frame.getContentPane().add(scrollPane);
		
		
		scrollPane.setViewportView(lstProject);
		
		JButton btnExit = new JButton("離開");
		btnExit.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				System.exit(0);
			}
		});
		btnExit.setBounds(433, 289, 87, 39);
		frame.getContentPane().add(btnExit);
		
		btnLoadTemplate = new JButton("1. 讀取空白標籤Excel");
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
		btnLoadTemplate.setBounds(10, 21, 158, 40);
		frame.getContentPane().add(btnLoadTemplate);
		
		JLabel label_3 = new JLabel("規格:");
		label_3.setBounds(175, 46, 36, 15);
		frame.getContentPane().add(label_3);
		
		lblFormat = new JLabel("");
		lblFormat.setForeground(Color.MAGENTA);
		lblFormat.setFont(new Font("PMingLiU", Font.BOLD, 12));
		lblFormat.setBounds(210, 46, 85, 15);
		frame.getContentPane().add(lblFormat);
		
		JLabel label_4 = new JLabel("\u6A94\u6848");
		label_4.setBounds(175, 27, 36, 15);
		frame.getContentPane().add(label_4);
		
		txfBlankPath = new JTextField();
		txfBlankPath.setColumns(10);
		txfBlankPath.setBounds(210, 23, 314, 21);
		frame.getContentPane().add(txfBlankPath);
		
		
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
        	
        	stList = new ArrayList<StickerClass>();
        	
        	Cell cell = null;
        	
        	ro_sheet = ro_data_workbook.getSheet(0);
        	
        	String sno = ""; //系統號
        	String snoa = ""; //縮號
        	String pcount = "";//株數
        	String src = ""; //來源
        	String note = ""; //備註
        	
        	int rows = ro_sheet.getRows();
        	int cols = ro_sheet.getColumns();
        	
        	if (cols > TEMP_COLS){
        		lblMsg.setText("內容格式錯誤,偵測到"+ cols + "個欄位; (資料欄位不得超過"+TEMP_COLS+"欄)");
        		return;
        	}
        	proj_count=0;
        	for (int r=2;r<rows;r++){
        		
        		//read rows
        		Cell crow[] = ro_sheet.getRow(r);
        		sno = crow[0].getContents();
        		snoa = crow[1].getContents();
        		pcount = crow[2].getContents();
        		src = crow[3].getContents();
        		note = crow[4].getContents();
        		        		
        		if (sno.length() != 0){
        			proj_count++;
        			listModel.addElement(proj_count+": " +sno+" ; " + snoa + " ; " + pcount + " ; " + src + " ; " + note);
        			stList.add(new StickerClass(sno, snoa, pcount, src, note));
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
	
	public class StickerClass{
    	public String sno = ""; //系統號
    	public String snoa = ""; //縮號
    	public String pcount = "";//株數
    	public String src = ""; //來源
    	public String note = ""; //備註
    	
    	public StickerClass(String _sno, String _snoa, String _pcount, String _src, String _note){
    		this.sno = _sno;
    		this.snoa = _snoa;
    		this.pcount = _pcount;
    		this.src = _src;
    		this.note = _note;
    	}
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
		
		StickerClass st = null;
		String brk = "\n";
		for (int pc=1; pc<=proj_count; pc++){
			//String data = "abc"+"\012"+"def"+"\012"+"ghi"+"\012";
			st = stList.get(pc-1);
			String data = st.sno + brk + st.snoa + brk + st.pcount + brk + st.src + brk + st.note; 

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
				System.out.println("it's empty");
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
