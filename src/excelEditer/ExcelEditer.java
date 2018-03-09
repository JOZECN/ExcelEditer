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
	private static JLabel ifLabel = new JLabel("��Դ��� ��");
	private static JLabel ifLabel2 = new JLabel("��Դ��ʽ ��");
	private static JLabel ofLabel = new JLabel("Ŀ���� ��");
	private static JLabel ofLabel2 = new JLabel("Ŀ���ʽ ��");
	private static JButton ifButton = new JButton("ѡ��");
	private static JButton ofButton = new JButton("ѡ��");
	private static JButton wButton = new JButton("д��");
	private static JRadioButton rb360 = new JRadioButton("360");
	private static JRadioButton rbshenma = new JRadioButton("����");
	private static JRadioButton rbsougou = new JRadioButton("�ѹ�");
	private static TextField ifText = new TextField();
	private static TextField ofText = new TextField();

	public ExcelEditer() {
		ifLabel.setFont(new Font("����", Font.BOLD, 12));
		ifLabel.setBounds(20, 20, 100, 30);
		getContentPane().add(ifLabel);
		
		ifText.setText("��ѡ���ת��excel�ļ�...");
		ifText.setBackground(Color.WHITE);
		ifText.setEnabled(false);
		ifText.setEditable(false);
		ifText.setForeground(Color.BLACK);
		ifText.setFont(new Font("����", Font.PLAIN, 14));
		ifText.setBounds(120, 20, 270, 30);
		getContentPane().add(ifText);
		
		ifButton.setFont(new Font("����", Font.BOLD, 12));
		ifButton.setBounds(400, 20, 80, 30);
		getContentPane().add(ifButton);
		
		ifLabel2.setFont(new Font("����", Font.BOLD, 12));
		ifLabel2.setBounds(20, 60, 100, 30);
		getContentPane().add(ifLabel2);
		
		rb360.setFont(new Font("����", Font.BOLD, 12));
		rb360.setBounds(120, 60, 60, 30);
		getContentPane().add(rb360);
		
		rbshenma.setFont(new Font("����", Font.BOLD, 12));
		rbshenma.setBounds(200, 60, 60, 30);
		getContentPane().add(rbshenma);
		
		rbsougou.setFont(new Font("����", Font.BOLD, 12));
		rbsougou.setBounds(280, 60, 60, 30);
		getContentPane().add(rbsougou);
		
		
		ofLabel.setFont(new Font("����", Font.BOLD, 12));
		ofLabel.setBounds(20, 140, 100, 30);
		getContentPane().add(ofLabel);
		
		ofText.setText("��ѡ���д��excel�ļ�...");
		ofText.setBackground(Color.WHITE);
		ofText.setEnabled(false);
		ofText.setEditable(false);
		ofText.setForeground(Color.BLACK);
		ofText.setFont(new Font("����", Font.PLAIN, 14));
		ofText.setBounds(120, 140, 270, 30);
		getContentPane().add(ofText);
		
		ofButton.setFont(new Font("����", Font.BOLD, 12));
		ofButton.setBounds(400, 140, 80, 30);
		getContentPane().add(ofButton);
		
		ofLabel2.setFont(new Font("����", Font.BOLD, 12));
		ofLabel2.setBounds(20, 180, 100, 30);
		getContentPane().add(ofLabel2);
		
		
		wButton.setFont(new Font("����", Font.BOLD, 12));
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
		JFileChooser chooser = new JFileChooser(); //����ѡ���ļ�����
		chooser.setDialogTitle("��ѡ���ļ�");//���ñ���
		chooser.setMultiSelectionEnabled(false);
		chooser.setFileSelectionMode(0);//0��ʾֻ��ѡ���ļ���1��ʾֻ��ѡ���ļ��У�2��ʾ����ѡ
		FileNameExtensionFilter filter = new FileNameExtensionFilter("Excel(*xls)","xls");//�����ѡ���ļ�����
		chooser.setFileFilter(filter); //���ÿ�ѡ���ļ�����
		int returnVal = chooser.showOpenDialog(null); //��ѡ���ļ��Ի���,null������Ϊ�㵱ǰ�Ĵ���JFrame��Frame
	    if(JFileChooser.APPROVE_OPTION == returnVal){
	    	ifile = chooser.getSelectedFile(); //fileΪ�û�ѡ���excel
		    String filepath = chooser.getSelectedFile().getAbsolutePath(); //��ȡ����·��  
		    ifText.setText(filepath);
		    //System.out.println(filepath);
	    }
    }
	
	public static void ofButtonMethod(){
		JFileChooser chooser = new JFileChooser(); //����ѡ���ļ�����
		chooser.setDialogTitle("��ѡ���ļ�");//���ñ���
		chooser.setMultiSelectionEnabled(false);
		chooser.setFileSelectionMode(0);//0��ʾֻ��ѡ���ļ���1��ʾֻ��ѡ���ļ��У�2��ʾ����ѡ
		FileNameExtensionFilter filter = new FileNameExtensionFilter("Excel(*xls)","xls");//�����ѡ���ļ�����
		chooser.setFileFilter(filter); //���ÿ�ѡ���ļ�����
		int returnVal = chooser.showOpenDialog(null); //��ѡ���ļ��Ի���,null������Ϊ�㵱ǰ�Ĵ���JFrame��Frame
	    if(JFileChooser.APPROVE_OPTION == returnVal){
	    	ofile = chooser.getSelectedFile(); //fileΪ�û�ѡ���excel
		    String filepath = chooser.getSelectedFile().getAbsolutePath(); //��ȡ����·��  
		    ofText.setText(filepath);
		    //System.out.println(filepath);
	    }
    }
	
	public void readExcel(File file) {  
        try {
            // ��������������ȡExcel
            InputStream is = new FileInputStream(file.getAbsolutePath());
            // jxl�ṩ��Workbook�� 
            Workbook wb = Workbook.getWorkbook(is);
            // Excel��ҳǩ���� 
            Sheet[] sheets = wb.getSheets();
            if (sheets != null){
            	for (Sheet sheet : sheets){
            		// �������
            		int rows = sheet.getRows();
            		// �������
            		int cols = sheet.getColumns();
            		// ��ȡ����
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
                // ÿ��ҳǩ����һ��Sheet����  
                Sheet sheet = wb.getSheet(index);  
                // sheet.getRows()���ظ�ҳ��������
                for (int i = 0; i < sheet.getRows(); i++) {  
                    // sheet.getColumns()���ظ�ҳ��������  
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
		          sheet.addCell(new Label(col, row, "־��" + row + col));
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
