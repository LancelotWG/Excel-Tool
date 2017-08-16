import java.awt.FlowLayout;
import java.awt.GridLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.lang.reflect.InvocationTargetException;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JTextField;
import javax.swing.SwingUtilities;
import javax.swing.filechooser.FileFilter;

import org.apache.log4j.Logger;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;

public  class  Main implements ActionListener{
	 
	static String basic="";
	static String abnormal="";
	/*static String delay="";
	static String delete="";*/
	static String postSwapInput =null; // post swap inputs
	static String result="";
	static String ref="";
	//Broken through flight 
	static String brokenResult = "";
	
	static JFrame frame;
	static JTextField basictext=new JTextField(20);
	static JTextField abnormaltext=new JTextField(20);
	/*static JTextField delaytext=new JTextField(20);
	static JTextField swaptext=new JTextField(20);
	static JTextField deletetext=new JTextField(20);*/
	static JTextField resulttext=new JTextField(20);
	static JTextField reftext=new JTextField(20);
	
	static JButton basicbt=new JButton("Choose");
	static JButton abnormalbt=new JButton("Choose");
	/*static JButton delaybt=new JButton("Choose");
	static JButton swapbt=new JButton("Choose");
	static JButton deletebt=new JButton("Choose");*/
	static JButton  resultbt=new JButton("Choose");
	static JButton  refbt=new JButton("Choose");
	
	
	
	static JButton calbt=new JButton("Calculate");
	static JButton reset=new JButton("Reset");
	
	static boolean gui=true;
	private static Logger logger = Logger.getLogger(Main.class);  
	public static void main(String args[])
	{     
			
		if(args.length==0)
			GUI();
		else
			CMD(args);
	}
	
	public static void CMD(String args[])
	{		
		gui=false;
		/*for(int i=0;i<args.length;i++)
			System.out.println(args[i]);*/
		if(parseArgs(args))
		{
			if(basic.equals("") ||result.equals("") ||ref.equals(""))
			{
				System.out.println("the basic(or result,or relation) file is null");
				return;
			}
			
			try {
				ReadExcel.handle(basic, abnormal, result,ref,postSwapInput,brokenResult);
			//	ReadExcel.handle(basic, delay, delete, swap, result,ref);
			} catch (InvocationTargetException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (InterruptedException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
		else
		{
			System.out.println("The args format is wrong!");
			printHelp();
		}
	}
	
	public static boolean parseArgs(String args[])
	{
		String entry;
		char flag;
		if(args.length==1 && (args[0].equals("-help") || args[0].equals("-Help") ||args[0].equals("-H") ||args[0].equals("-h")) )
		{
			//printHelp();
			return false;
		}
		for(int i=0;i<args.length;i++)
		{
			entry=args[i].trim();
			if(entry.charAt(1)=='=' &&(entry.charAt(0)=='B' || entry.charAt(0)=='b' ||
					entry.charAt(0)=='A' ||  entry.charAt(0)=='a' || 					
					entry.charAt(0)=='O' ||  entry.charAt(0)=='o' ||
					entry.charAt(0)=='R' ||  entry.charAt(0)=='r' || 
					entry.charAt(0)=='S' ||  entry.charAt(0)=='s' 
					/*修改自兰望桂*/
					|| entry.charAt(0)=='T' || entry.charAt(0)=='t'))
				/*entry.charAt(0)=='C' ||  entry.charAt(0)=='c' ||
				entry.charAt(0)=='S' ||  entry.charAt(0)=='s' ||*/
			{
				flag=entry.charAt(0);
				if(flag=='B' ||flag=='b' )
				{					
					basic=entry.substring(2);
					if(!isCorrectFile(basic))
						return false;
				}
				else if(flag=='A' ||flag=='a')
				{
					abnormal=entry.substring(2);
					if(!isCorrectFile(abnormal))
						return false;
				}
				/*else if(flag=='D'  ||flag=='d')
				{
					delay=entry.substring(2);
					if(!isCorrectFile(delay))
						return false;
				}
				else if(flag=='C' ||flag=='c')
				{
					delete=entry.substring(2);
					if(!isCorrectFile(delete))
						return false;
				}*/
				else if(flag=='S' ||flag=='s')
				{
					postSwapInput=entry.substring(2);
					if(!isCorrectFile(postSwapInput))
						return false;
				}
				else if(flag=='O'||flag=='o')
				{
					result=entry.substring(2);
					if(!isCorrectFile(result))
						return false;
				}
				else if(flag=='R'||flag=='r')
				{
					ref=entry.substring(2);	
					if(!isCorrectFile(ref))
						return false;
				}
				//修改自兰望桂
				/**
				 * 增加联程航班中断输出文件
				 */
				else if(flag=='T'||flag=='t')
				{
					brokenResult=entry.substring(2);	
					if(!isCorrectFile(brokenResult))
						return false;
				}
				//修改自兰望桂
			}
			else 
			{
				System.out.println("this program only accept file args:B/b,A/a,O/o,R/r,T/t,and the format is like B=File_Path");
				return false;
			}
					
		}
		return true;
	}
	
	public static boolean isCorrectFile(String name)
	{
	
		if(name.length()<6)//at least X.xlsx
		{
			System.out.println("Are you adding the file format suffix,cause the name("+name+") is too short.you should input like basicdata.xlsx,rather than basicdata");
			return false;
		}
		
		if(name.charAt(0)=='"' )
		{
			if(name.charAt(name.length()-1)=='"')
				name=name.substring(1, name.length()-2);
			else
			{
				System.out.println("There is missing a right quote in the file path("+name+")");
				return false;
			}
		}
	//if(!name.toLowerCase().endsWith(".xlsx")  )//|| name.toLowerCase().endsWith(".xlsx")
	//		return false;
		File file=new File(name);
		if(!file.exists())
		{
			System.out.println("File("+name+") doesn't exist,please check it aggain and retry!");
			return false;
		}

		return true;
	}
	public static void printHelp()
	{
		System.out.println("************************************************************************");
		System.out.println("************************************************************************");
		System.out.println("************************************************************************");
		System.out.println("The cmd is like : R=reflection_file_dir  O=result_file dir  B=Basic_file_dir "
				+ "S=Swap_file_dir  D=Delay_file dir  C=Cancel_file_dir T=Borken_file_dir\n");
		System.out.println("\nIf the dir has blank,then quotes the dir,for example B=\"Basic data.xlsx\"\n");
		System.out.println("\nMake sure the file is exisit and is correct format,and the args order is not relative!\n");
		System.out.println("\nPlus:the program only accept xlsx file!\n");
		System.out.println("For help:use -help or -h");
		System.out.println("************************************************************************");
		System.out.println("************************************************************************");
	}
	public static void GUI()
	{
		gui=true;
		frame =new JFrame("Excel Operate");
		frame.setSize(450,250);
		frame.setLocationRelativeTo(null); 
		frame.setDefaultCloseOperation(frame.DISPOSE_ON_CLOSE);
		
		JPanel p = new JPanel();
		JPanel p1 = new JPanel();
		JPanel p2 = new JPanel();
		/*JPanel p3 = new JPanel();
		JPanel p4 = new JPanel();*/
		JPanel p5 = new JPanel();
		JPanel p6 = new JPanel();
		JPanel p7 = new JPanel();
		
		FlowLayout fl=new FlowLayout();
		p1.setLayout(fl);
		p2.setLayout(fl);
		/*p3.setLayout(fl);
		p4.setLayout(fl);*/
		p5.setLayout(fl);
		p6.setLayout(fl);
		p7.setLayout(fl);
		
		
		JLabel basiclabel=new JLabel("Basic          file:     ");
		JLabel abnormallabel=new JLabel("Abnormal   file:     ");
		/*JLabel deletelabel=new JLabel("Cancel         file:     ");
		JLabel delaylabel=new JLabel("Delay          file:     ");
		JLabel swaplabel=new JLabel("Swap          file:     ");		*/
		JLabel resultlabel=new JLabel("Result        file:     ");
		JLabel reflabel=new JLabel("Relation scheme:");
		
		basictext.setEditable(false);
		abnormaltext.setEditable(false);
		resulttext.setEditable(false);
		reftext.setEditable(false);
		
		p1.add(basiclabel);
		p1.add(basictext);
		p1.add(basicbt);
		
		p2.add(abnormallabel);
		p2.add(abnormaltext);
		p2.add(abnormalbt);
		
		/*p3.add(swaplabel);
		p3.add(swaptext);
		p3.add(swapbt);
		
		p4.add(deletelabel);
		p4.add(deletetext);
		p4.add(deletebt);*/
		
		p5.add(resultlabel);
		p5.add(resulttext);
		p5.add(resultbt);
		
		p6.add(reflabel);
		p6.add(reftext);
		p6.add(refbt);
		
		p7.add(calbt);
		p7.add(reset);
		
		Main main=new Main();
		basicbt.addActionListener(main);
		abnormalbt.addActionListener(main);
		/*swapbt.addActionListener(main);
		delaybt.addActionListener(main);
		deletebt.addActionListener(main);*/
		resultbt.addActionListener(main);
		refbt.addActionListener(main);
		
		calbt.addActionListener(main);
		reset.addActionListener(main);
		//p.setLayout(new GridLayout(7,1));
		p.setLayout(new GridLayout(5,1));
		p.add(p1);//basic
	//	p.add(p4);//cancel
		p.add(p2);//delay
	//	p.add(p3);	//swap	
		p.add(p5);//resulf
		p.add(p6); //ref
		p.add(p7);//button
		
		frame.setContentPane(p);
		frame.setVisible(true);
	}

	@Override
	public void actionPerformed(ActionEvent e) {
		// TODO Auto-generated method stub
		if(e.getSource()==reset)
		{
			basictext.setText("");	
			abnormaltext.setText("");
		/*	delaytext.setText("");	
			deletetext.setText("");	
			swaptext.setText("");	*/
			reftext.setText("");
			resulttext.setText("");	
			
			basic="";
			abnormal="";
			/*delay="";
			swap="";
			delete="";*/
			result="";
			ref="";
			
		}
	    else if(e.getSource()==basicbt)
		{
			basic=getFile();
			basictext.setText(basic);			
		}
	    else if(e.getSource()==abnormalbt)
		{
	    	abnormal=getFile();
	    	abnormaltext.setText(abnormal);
		}
		/*else if(e.getSource()==delaybt)
		{
			delay=getFile();
			delaytext.setText(delay);
		}
		else if(e.getSource()==deletebt)
		{
			delete=getFile();
			deletetext.setText(delete);
		}
		else if(e.getSource()==swapbt)
		{
			swap=getFile();
			swaptext.setText(swap);
		}*/
		else if(e.getSource()==resultbt)
		{
			result=getFile();
			resulttext.setText(result);
		}
		else if(e.getSource()==refbt)
		{
			ref=getFile();
			reftext.setText(ref);
		}
		else
		{
			
			if(basic.equals("") ||result.equals("") ||ref.equals(""))
			{
				JOptionPane.showMessageDialog(frame.getContentPane(),
						 "the basic(or result,or relation) file is null", "Error", JOptionPane.INFORMATION_MESSAGE);
				return;
			}
			
			
		/*	basicbt.setEnabled(false);				
			deletebt.setEnabled(false);
			delaybt.setEnabled(false);
			swapbt.setEnabled(false);
			resultbt.setEnabled(false);
			refbt.setEnabled(false);
			
			calbt.setEnabled(false);
			reset.setEnabled(false);*/
			
			
    		/*System.out.println( basic);
    		System.out.println( delete);
    		System.out.println( delay);
    		System.out.println( swap);
			System.out.println( result);
			System.out.println( ref);*/
			
			/*Cal cal=new Cal(basic, delay, delete, swap, result,ref);
			Thread thread =new Thread(cal);
		    thread.start();*/
			
			/*frame.setVisible(false);
			ShowPanel panel=new ShowPanel(frame);*/
			
			frame.dispose();
			frame=new ShowPanel();
			frame.invalidate();	
			Thread t=new Thread(new Cal(basic, abnormal, result,ref, postSwapInput, brokenResult));
			//Thread t=new Thread(new Cal(basic, delay, delete, swap, result,ref));
			t.start();
			//SwingUtilities.invokeLater();
		
			frame.invalidate();
			/* calbt.setEnabled(true);
			 basicbt.setEnabled(true);
			deletebt.setEnabled(true);
			delaybt.setEnabled(true);
			swapbt.setEnabled(true);
			resultbt.setEnabled(true);
			refbt.setEnabled(true);*/
		}
	}
	
	public void printInfo(ReadExcel.InfoLevel level,String info)
	{
		if(level==ReadExcel.InfoLevel.Info)
			logger.info(info);
		else if(level==ReadExcel.InfoLevel.Warn)
			logger.warn(info);
		else
			logger.error(info);
		
		String val=level.getName()+info;
		if(!Main.gui)
		{
			System.out.println(val);
		}else
		{
			JOptionPane.showMessageDialog(frame.getContentPane(),
					 info, level.getName(), JOptionPane.INFORMATION_MESSAGE);
		}
	
	}
	public  String getFile()
	{
		JFileChooser jf = new JFileChooser(".");  
		
		ExcelFileFilter excelFilter = new ExcelFileFilter(); 
		jf.addChoosableFileFilter(excelFilter);  
		jf.setFileFilter(excelFilter);  
		  
		jf.setFileSelectionMode(JFileChooser.FILES_ONLY);  
		if(jf.showDialog(null,null)!=JFileChooser.APPROVE_OPTION)
			return "";
	
		File fi = jf.getSelectedFile(); 
		String filename=fi.getAbsolutePath();
	    if(!fi.exists() )
	    {
	    	printInfo(ReadExcel.InfoLevel.Error, "File:"+filename+" is not exist!");
	    	return "";
	    }
	    if(!filename.endsWith(".xlsx"))
	    {
	    	printInfo(ReadExcel.InfoLevel.Error, "File:"+filename+" format is not xlsx,this program only accept xlsx file!");
	    	return "";
	    }
	   
		return fi.getAbsolutePath();
	}
	
}


class ExcelFileFilter extends FileFilter {    
    public String getDescription() {    
        return "*.xls;*.xlsx";    
    }    
    
    public boolean accept(File file) {    
        String name = file.getName();    
        return file.isDirectory() || name.toLowerCase().endsWith(".xlsx"); //name.toLowerCase().endsWith(".xls") || 
    }    
}  

class  Cal implements Runnable{

	String basic;
	String abnormal;
	static String postSwapInput =""; // post swap inputs
	String brokenResult;
	/*String delay;
	String delete;
	String swap;*/
	String result;
	String ref;
	public Cal(String basic,String abnormal,String result,String ref, String postSwapInput, String brokenResult)//String delay,String delete,String swap,
	{
		this.basic=basic;
		this.abnormal=abnormal;
	/*	this.delay=delay;
		this.delete=delete;
		this.swap=swap;*/
		this.result=result;
		this.ref=ref;
		this.postSwapInput = postSwapInput;
		this.brokenResult = brokenResult;
	}
	@Override
	public void run() {
		
		
		try {
			//ReadExcel.handle(basic, delay, delete, swap, result,ref);
			ReadExcel.handle(basic,abnormal, result,ref, postSwapInput, brokenResult);// delay, delete, swap,
		} catch (InvocationTargetException e) {
		
			e.printStackTrace();
		} catch (InterruptedException e) {		
			e.printStackTrace();
		}
		
	}
	
}


