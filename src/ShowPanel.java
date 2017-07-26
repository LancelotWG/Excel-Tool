import java.awt.GridLayout;
import java.awt.TextArea;

import javax.swing.BorderFactory;

import javax.swing.JFrame;

import javax.swing.JPanel;
import javax.swing.JScrollPane;

import javax.swing.border.TitledBorder;

public class ShowPanel extends JFrame{
	
	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;
	TextArea infota;
	TextArea errorta;
	
	//String basic,String delay,String delete,String swap,String result,String ref
	public ShowPanel()
	{
		
		 super("Handling Excel data...");
		
		 
		 JPanel panel=new JPanel();
	     this.setContentPane(panel);
	     this.setLocationRelativeTo(null);
	     	 	   
	    
	     infota=new TextArea(12,70);
	     JScrollPane scrolli = new JScrollPane(infota);  
	     scrolli.setHorizontalScrollBarPolicy( JScrollPane.HORIZONTAL_SCROLLBAR_AS_NEEDED);  
	     scrolli.setVerticalScrollBarPolicy( JScrollPane.VERTICAL_SCROLLBAR_AS_NEEDED); 	    
	     TitledBorder borderinfo=BorderFactory.createTitledBorder(BorderFactory.createLoweredBevelBorder());
	     borderinfo.setTitle("Progress info:");
	     scrolli.setBorder(borderinfo);
	    
	     errorta=new TextArea(12,70);
	     JScrollPane scrolle = new JScrollPane(errorta);  
	     scrolle.setHorizontalScrollBarPolicy( JScrollPane.HORIZONTAL_SCROLLBAR_AS_NEEDED);  
	     scrolle.setVerticalScrollBarPolicy( JScrollPane.VERTICAL_SCROLLBAR_AS_NEEDED); 	    
	     TitledBorder bordererror=BorderFactory.createTitledBorder(BorderFactory.createLoweredBevelBorder());
	     bordererror.setTitle("Process erro:");
	     scrolle.setBorder(bordererror);
	     
	     JPanel up=new JPanel();
	     up.setLayout(new GridLayout(2,1));
	     up.add(scrolli);
	     up.add(scrolle);
	     
	     panel.add(up);
	     
	     this.setSize(580,500);//length,width
	     
	    // this.setSize(780,600);
	     this.setDefaultCloseOperation(this.DISPOSE_ON_CLOSE);
	     this.setVisible(true);
	   //  System.out.println("  ------");	    	   	    	     
	     
	     
	}
	
	
	
	
}
