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
	private static JLabel ifLabel = new JLabel("Դ��� ��");
	private static JButton ifButton = new JButton("����");
	private static JRadioButton rb360 = new JRadioButton("360");
	private static JRadioButton rbshenma = new JRadioButton("����");
	private static JRadioButton rbsougou = new JRadioButton("�ѹ�");
	private static TextField ifText = new TextField();

	public ExcelEditer() {
		ifLabel.setFont(new Font("����", Font.BOLD, 12));
		ifLabel.setBounds(20, 20, 60, 30);
		getContentPane().add(ifLabel);
		
		ifButton.setFont(new Font("����", Font.BOLD, 12));
		ifButton.setBounds(400, 20, 80, 30);
		getContentPane().add(ifButton);
		
		rb360.setFont(new Font("����", Font.BOLD, 12));
		rb360.setBounds(20, 60, 100, 30);
		getContentPane().add(rb360);
		
		rbshenma.setFont(new Font("����", Font.BOLD, 12));
		rbshenma.setBounds(140, 60, 100, 30);
		getContentPane().add(rbshenma);
		
		rbsougou.setFont(new Font("����", Font.BOLD, 12));
		rbsougou.setBounds(260, 60, 100, 30);
		getContentPane().add(rbsougou);
		
		ifText.setText("��ѡ���ת��excel�ļ�...");
		ifText.setBackground(Color.WHITE);
		ifText.setEnabled(false);
		ifText.setEditable(false);
		ifText.setForeground(Color.BLACK);
		ifText.setFont(new Font("����", Font.PLAIN, 14));
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
		JFileChooser chooser = new JFileChooser(); //����ѡ���ļ�����
		chooser.setDialogTitle("��ѡ���ļ�");//���ñ���
		chooser.setMultiSelectionEnabled(false);
		chooser.setFileSelectionMode(0);//0��ʾֻ��ѡ���ļ���1��ʾֻ��ѡ���ļ��У�2��ʾ����ѡ
		FileNameExtensionFilter filter = new FileNameExtensionFilter("Excel(*xls)","xls");//�����ѡ���ļ�����
		chooser.setFileFilter(filter); //���ÿ�ѡ���ļ�����
		int returnVal = chooser.showOpenDialog(null); //��ѡ���ļ��Ի���,null������Ϊ�㵱ǰ�Ĵ���JFrame��Frame
	    if(JFileChooser.APPROVE_OPTION == returnVal){
	    	//file = chooser.getSelectedFile(); //fileΪ�û�ѡ���ͼƬ�ļ�
		    String filepath = chooser.getSelectedFile().getAbsolutePath(); //��ȡ����·��  
		    ifText.setText(filepath);
		    System.out.println(filepath);
	    }
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
