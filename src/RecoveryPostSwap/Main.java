package RecoveryPostSwap;

import javax.swing.text.html.HTMLDocument.Iterator;

public class Main {

	public static void main(String[] args)
	{
		String filePath = "./postRecovery-test1.txt";
		//String filePath = "./typhoon20170324soln6-2.txt";
		SwapsPaser paser = new SwapsPaser();
		paser.paserData(filePath);
		java.util.Iterator<SwapEntity> it =   paser.swapList.iterator(); 
		while(it.hasNext()){  
			SwapEntity item = it.next();
			System.out.println(item.swappedAricraft + " switches to the path of -> " + item.orgAircraft);  
			 
			}
		System.out.println(paser.swapList.size());
		System.out.println(paser.subfleetMap);
	}
}
