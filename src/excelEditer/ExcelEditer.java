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
import java.awt.TextField;
import java.awt.Font;
import java.awt.Color;

public class ExcelEditer extends JFrame implements ActionListener{
	private static File file;
	private static final int DEFAULT_WIDTH=500;
	private static final int DEFAULT_HEIGHT=300;
	private static JLabel ifLabel = new JLabel("源表格 ：");
	private static JButton ifButton = new JButton("导入");
	private static JRadioButton rb360 = new JRadioButton("360");
	private static JRadioButton rbshenma = new JRadioButton("神马");
	private static JRadioButton rbsougou = new JRadioButton("搜狗");
	private static TextField ifText = new TextField();

	public ExcelEditer() {
		ifLabel.setFont(new Font("宋体", Font.BOLD, 12));
		ifLabel.setBounds(20, 20, 60, 30);
		getContentPane().add(ifLabel);
		
		ifButton.setFont(new Font("宋体", Font.BOLD, 12));
		ifButton.setBounds(400, 20, 80, 30);
		getContentPane().add(ifButton);
		
		rb360.setFont(new Font("宋体", Font.BOLD, 12));
		rb360.setBounds(20, 60, 100, 30);
		getContentPane().add(rb360);
		
		rbshenma.setFont(new Font("宋体", Font.BOLD, 12));
		rbshenma.setBounds(140, 60, 100, 30);
		getContentPane().add(rbshenma);
		
		rbsougou.setFont(new Font("宋体", Font.BOLD, 12));
		rbsougou.setBounds(260, 60, 100, 30);
		getContentPane().add(rbsougou);
		
		ifText.setText("请选择待转换excel文件...");
		ifText.setBackground(Color.WHITE);
		ifText.setEnabled(false);
		ifText.setEditable(false);
		ifText.setForeground(Color.BLACK);
		ifText.setFont(new Font("宋体", Font.PLAIN, 14));
		ifText.setBounds(80, 26, 300, 20);
		getContentPane().add(ifText);
		
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
	    	//file = chooser.getSelectedFile(); //file为用户选择的图片文件
		    String filepath = chooser.getSelectedFile().getAbsolutePath(); //获取绝对路径  
		    ifText.setText(filepath);
		    System.out.println(filepath);
	    }
    }
	
	public void readExcel(File file) {  
        try {  
            // 创建输入流，读取Excel  
            InputStream is = new FileInputStream(file.getAbsolutePath());  
            // jxl提供的Workbook类  
            Workbook wb = Workbook.getWorkbook(is);  
            // Excel的页签数量  
            int sheet_size = wb.getNumberOfSheets();  
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
            }  
        } catch (FileNotFoundException e) {  
            e.printStackTrace();
        } catch (BiffException e) {  
            e.printStackTrace();
        } catch (IOException e) {  
            e.printStackTrace();
        }  
    }
}
