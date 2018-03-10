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
	private static int sheet_size;
	private static int totalRow;
	private static int totalCol;
	private static int total;
	private int[][] sogouTo360 = {
		{1,1,1,2},
		{1,2,1,3}
	}; //页签，源列数，目的列数
	
	private static final int DEFAULT_WIDTH=500;
	private static final int DEFAULT_HEIGHT=300;
	private static JLabel ifLabel = new JLabel("来源表格 ：");
	private static JLabel ifLabel2 = new JLabel("来源格式 ：");
	private static JLabel ofLabel = new JLabel("目标表格 ：");
	private static JLabel ofLabel2 = new JLabel("目标格式 ：");
	private static JButton ifButton = new JButton("选择");
	private static JButton ofButton = new JButton("选择");
	private static JButton wButton = new JButton("写入");
	private static JLabel wLabel = new JLabel("");
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
		
		wLabel.setFont(new Font("宋体", Font.BOLD, 12));
		wLabel.setBounds(320, 220, 100, 30);
		getContentPane().add(wLabel);
		
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
            	wLabel.setForeground(Color.RED);
            	wLabel.setText("写入中...");
            	
            	obj.readExcel(ifile);
            	//obj.writeExcel(ofile);
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
            InputStream is = new FileInputStream(file.getAbsolutePath()); // 创建输入流，读取Excel
            Workbook wb = Workbook.getWorkbook(is); // jxl提供的Workbook类 
            sheet_size = wb.getNumberOfSheets(); // Excel的页签数量 
            for (int index = 0; index < sheet_size; index++) {  
                Sheet sheet = wb.getSheet(index); // 每个页签创建一个Sheet对象  
                totalRow = sheet.getRows(); // 获得行数
                totalCol = sheet.getColumns(); // 获得列数
                total = totalRow * totalCol; // 获得表格单元格总数
                System.out.println("当前：一共"+totalRow+"行，"+totalCol+"列，总计："+total);
                WritableWorkbook wb2 = Workbook.createWorkbook(ofile);
                WritableSheet sheet2 = wb2.createSheet("sheet1", 0);
                for (int col = 0; col < totalCol; col++) {
                    for (int row = 0; row < totalRow; row++) {
                        String cellinfo = sheet.getCell(col, row).getContents();
                        //if(cellinfo != null && !cellinfo.equals("")){
                        	writeExcel(wb2,sheet2,sheet_size,row,col,cellinfo);
                        //}
                    }  
                }  
            }
            wb.close();
        } catch (FileNotFoundException e) {  
            e.printStackTrace();
        } catch (BiffException e) {  
            e.printStackTrace();
        } catch (IOException e) {  
            e.printStackTrace();
        }
    }
	
	public void writeExcel(WritableWorkbook wb,WritableSheet sheet,int sheet_size,int row,int col,String cellinfo) {
		try{
		    System.out.println("当前："+(row+1)+"行，"+(col+1)+"列");
		    System.out.println(cellinfo);
		    sheet.addCell(new Label(col, row, cellinfo)); //写入表格单元格数据
		    if((col+1)*(row+1) == total){
		    	wb.write();
		    	wb.close();
		    	//System.out.println("done");
		    	wLabel.setForeground(Color.GREEN);
            	wLabel.setText("写入完成！");
		    }
		} catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (WriteException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
