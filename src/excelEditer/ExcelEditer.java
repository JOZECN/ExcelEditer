package excelEditer;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JButton;
import javax.swing.JRadioButton;
import javax.swing.JFileChooser;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

import java.awt.TextField;
import java.awt.Font;
import java.awt.Color;

public class ExcelEditer extends JFrame implements ActionListener{
	private static File ifile;
	private static File ofile;
	private static final int DEFAULT_WIDTH=500;
	private static final int DEFAULT_HEIGHT=300;
	private static JLabel ifLabel = new JLabel("来源表格 ：");
	private static JLabel ifLabel2 = new JLabel("来源格式 ：");
	private static JLabel ofLabel = new JLabel("目标表格 ：");
	private static JLabel ofLabel2 = new JLabel("目标格式 ：");
	private static JButton ifButton = new JButton("选择");
	private static JButton ofButton = new JButton("选择");
	private static JButton wButton = new JButton("写入");
	private static JRadioButton rb360 = new JRadioButton("360");
	private static JRadioButton rbshenma = new JRadioButton("神马");
	private static JRadioButton rbsougou = new JRadioButton("搜狗");
	private static TextField ifText = new TextField();
	private static TextField ofText = new TextField();

	public ExcelEditer() {
		ifLabel.setFont(new Font("宋体", Font.BOLD, 12));
		ifLabel.setBounds(20, 20, 100, 30);
		getContentPane().add(ifLabel);
		
		ifText.setText("请选择待转换excel文件...");
		ifText.setBackground(Color.WHITE);
		ifText.setEnabled(false);
		ifText.setEditable(false);
		ifText.setForeground(Color.BLACK);
		ifText.setFont(new Font("宋体", Font.PLAIN, 14));
		ifText.setBounds(120, 20, 270, 30);
		getContentPane().add(ifText);
		
		ifButton.setFont(new Font("宋体", Font.BOLD, 12));
		ifButton.setBounds(400, 20, 80, 30);
		getContentPane().add(ifButton);
		
		ifLabel2.setFont(new Font("宋体", Font.BOLD, 12));
		ifLabel2.setBounds(20, 60, 100, 30);
		getContentPane().add(ifLabel2);
		
		rb360.setFont(new Font("宋体", Font.BOLD, 12));
		rb360.setBounds(120, 60, 60, 30);
		getContentPane().add(rb360);
		
		rbshenma.setFont(new Font("宋体", Font.BOLD, 12));
		rbshenma.setBounds(200, 60, 60, 30);
		getContentPane().add(rbshenma);
		
		rbsougou.setFont(new Font("宋体", Font.BOLD, 12));
		rbsougou.setBounds(280, 60, 60, 30);
		getContentPane().add(rbsougou);
		
		
		ofLabel.setFont(new Font("宋体", Font.BOLD, 12));
		ofLabel.setBounds(20, 140, 100, 30);
		getContentPane().add(ofLabel);
		
		ofText.setText("请选择待写入excel文件...");
		ofText.setBackground(Color.WHITE);
		ofText.setEnabled(false);
		ofText.setEditable(false);
		ofText.setForeground(Color.BLACK);
		ofText.setFont(new Font("宋体", Font.PLAIN, 14));
		ofText.setBounds(120, 140, 270, 30);
		getContentPane().add(ofText);
		
		ofButton.setFont(new Font("宋体", Font.BOLD, 12));
		ofButton.setBounds(400, 140, 80, 30);
		getContentPane().add(ofButton);
		
		ofLabel2.setFont(new Font("宋体", Font.BOLD, 12));
		ofLabel2.setBounds(20, 180, 100, 30);
		getContentPane().add(ofLabel2);
		
		
		wButton.setFont(new Font("宋体", Font.BOLD, 12));
		wButton.setBounds(200, 220, 100, 30);
		getContentPane().add(wButton);
		
		getContentPane().setLayout(null);
		this.setTitle("Excel Editer By Jozecn");
		this.setSize(DEFAULT_WIDTH, DEFAULT_HEIGHT);
		this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		this.setResizable(false);
		this.setVisible(true);
	}
	
	public void actionPerformed(ActionEvent e){
		
	}

	public static void main(String[] args) {
		ExcelEditer obj = new ExcelEditer();
		//obj.readExcel(file);
		ifButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
            	ifButtonMethod();
            }
        });
		ofButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
            	ofButtonMethod();
            }
        });
		wButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
            	//ofButtonMethod();
            	obj.readExcel(ifile);
            	obj.writeExcel(ofile);
            }
        });
	}
	
	public static void ifButtonMethod(){
		JFileChooser chooser = new JFileChooser(); //创建选择文件对象
		chooser.setDialogTitle("请选择文件");//设置标题
		chooser.setMultiSelectionEnabled(false);
		chooser.setFileSelectionMode(0);//0表示只能选择文件，1表示只能选择文件夹，2表示均可选
		FileNameExtensionFilter filter = new FileNameExtensionFilter("Excel(*xls)","xls");//定义可选择文件类型
		chooser.setFileFilter(filter); //设置可选择文件类型
		int returnVal = chooser.showOpenDialog(null); //打开选择文件对话框,null可设置为你当前的窗口JFrame或Frame
	    if(JFileChooser.APPROVE_OPTION == returnVal){
	    	ifile = chooser.getSelectedFile(); //file为用户选择的excel
		    String filepath = chooser.getSelectedFile().getAbsolutePath(); //获取绝对路径  
		    ifText.setText(filepath);
		    //System.out.println(filepath);
	    }
    }
	
	public static void ofButtonMethod(){
		JFileChooser chooser = new JFileChooser(); //创建选择文件对象
		chooser.setDialogTitle("请选择文件");//设置标题
		chooser.setMultiSelectionEnabled(false);
		chooser.setFileSelectionMode(0);//0表示只能选择文件，1表示只能选择文件夹，2表示均可选
		FileNameExtensionFilter filter = new FileNameExtensionFilter("Excel(*xls)","xls");//定义可选择文件类型
		chooser.setFileFilter(filter); //设置可选择文件类型
		int returnVal = chooser.showOpenDialog(null); //打开选择文件对话框,null可设置为你当前的窗口JFrame或Frame
	    if(JFileChooser.APPROVE_OPTION == returnVal){
	    	ofile = chooser.getSelectedFile(); //file为用户选择的excel
		    String filepath = chooser.getSelectedFile().getAbsolutePath(); //获取绝对路径  
		    ofText.setText(filepath);
		    //System.out.println(filepath);
	    }
    }
	
	public void readExcel(File file) {  
        try {
            // 创建输入流，读取Excel
            InputStream is = new FileInputStream(file.getAbsolutePath());
            // jxl提供的Workbook类 
            Workbook wb = Workbook.getWorkbook(is);
            // Excel的页签数量 
            Sheet[] sheets = wb.getSheets();
            if (sheets != null){
            	for (Sheet sheet : sheets){
            		// 获得行数
            		int rows = sheet.getRows();
            		// 获得列数
            		int cols = sheet.getColumns();
            		// 读取数据
            		for (int row = 0; row < rows; row++){
            			for (int col = 0; col < cols; col++){
            				System.out.printf("%10s", sheet.getCell(col, row).getContents());
            			}
            			System.out.println();
            		}
            	}
            }
            wb.close();
            /*int sheet_size = wb.getNumberOfSheets();  
            for (int index = 0; index < sheet_size; index++) {  
                // 每个页签创建一个Sheet对象  
                Sheet sheet = wb.getSheet(index);  
                // sheet.getRows()返回该页的总行数
                for (int i = 0; i < sheet.getRows(); i++) {  
                    // sheet.getColumns()返回该页的总列数  
                    for (int j = 0; j < sheet.getColumns(); j++) {  
                        String cellinfo = sheet.getCell(j, i).getContents();  
                        System.out.println(cellinfo);  
                    }  
                }  
            }*/
        } catch (FileNotFoundException e) {  
            e.printStackTrace();
        } catch (BiffException e) {  
            e.printStackTrace();
        } catch (IOException e) {  
            e.printStackTrace();
        }
    }
	
	public void writeExcel(File file) {
		try{
		    WritableWorkbook wb = Workbook.createWorkbook(file);
		    WritableSheet sheet = wb.createSheet("sheet1", 0);
		    for (int row = 1; row < 20; row++)
		    {
		       for (int col = 1; col < 20; col++)
		       {
		          sheet.addCell(new Label(col, row, "志鹏" + row + col));
		       }
		    }
		    wb.write();
		    wb.close();
		} catch (FileNotFoundException e) {  
            e.printStackTrace();
        } catch (WriteException e) {  
            e.printStackTrace();
        } catch (IOException e) {  
            e.printStackTrace();
        }
    }
}
