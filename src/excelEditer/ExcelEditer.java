package excelEditer;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.border.EmptyBorder;
import javax.swing.JLabel;
import javax.swing.JButton;
import javax.swing.JRadioButton;
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
	private static final int DEFAULT_WIDTH=500;
	private static final int DEFAULT_HEIGHT=300;

	public ExcelEditer() {
		JLabel ifLabel = new JLabel("Դ���");
		ifLabel.setFont(new Font("����", Font.BOLD, 12));
		ifLabel.setBounds(20, 20, 60, 30);
		this.add(ifLabel);
		
		JButton ifButton = new JButton("����");
		ifButton.setFont(new Font("����", Font.BOLD, 12));
		ifButton.setBounds(400, 20, 80, 30);
		this.add(ifButton);
		
		JRadioButton rb360 = new JRadioButton("360");
		rb360.setFont(new Font("����", Font.BOLD, 12));
		rb360.setBounds(20, 60, 100, 30);
		this.add(rb360);
		
		JRadioButton rbshenma = new JRadioButton("����");
		rbshenma.setFont(new Font("����", Font.BOLD, 12));
		rbshenma.setBounds(140, 60, 100, 30);
		this.add(rbshenma);
		
		JRadioButton rbsougou = new JRadioButton("�ѹ�");
		rbsougou.setFont(new Font("����", Font.BOLD, 12));
		rbsougou.setBounds(260, 60, 100, 30);
		this.add(rbsougou);
		
		TextField ifField = new TextField();
		ifField.setBackground(Color.WHITE);
		ifField.setEnabled(false);
		ifField.setEditable(false);
		ifField.setForeground(Color.BLACK);
		ifField.setFont(new Font("����", Font.PLAIN, 12));
		ifField.setBounds(80, 20, 300, 30);
		this.add(ifField);
		
		this.setLayout(null);
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
		// �˴�ΪExcel�ļ�·�� 
        File file = new File("D:/JxlTest.xls");
        obj.readExcel(file);
	}
	
	public void readExcel(File file) {  
        try {  
            // ��������������ȡExcel  
            InputStream is = new FileInputStream(file.getAbsolutePath());  
            // jxl�ṩ��Workbook��  
            Workbook wb = Workbook.getWorkbook(is);  
            // Excel��ҳǩ����  
            int sheet_size = wb.getNumberOfSheets();  
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
