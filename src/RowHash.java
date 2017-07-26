
public class RowHash implements Comparable<RowHash>{

	String fligth_num;
	String d_time;
	String departure;
	String arrival;
	public RowHash(String num,String time,String d,String a)
	{
		fligth_num=num;
		d_time=time;
		this.departure=d;
		this.arrival=a;;
	}
	
	
	@Override
	public String toString() {	
		return "Flight number :"+this.fligth_num+",Date :"+d_time+",From :"+departure+",To :"+arrival;
	}


	@Override
	public int compareTo(RowHash arg0) {
		//System.out.println(" equal arg0 f:"+arg0.fligth_num+" t:"+arg0.d_time);
		//System.out.println(" equal this f:"+this.fligth_num+" t:"+this.d_time);
		
		if(fligth_num.equals(arg0.fligth_num) && d_time.equals(arg0.d_time) 
				&& departure.equals(arg0.departure)  && arrival.equals(arg0.arrival))
		{//   System.out.println("true");
			return 0;
		}
		
		//System.out.println("false");
		return -1;
	}
	
	
	@Override
	public int hashCode() {
		return this.fligth_num.hashCode()+ this.d_time.hashCode()
	              +this.departure.hashCode()+ this.arrival.hashCode();
	}
	
	
	@Override
	public boolean equals(Object arg0) {
		RowHash other=(RowHash)arg0;
	//	System.out.println(" equal arg0 f:"+other.fligth_num+" t:"+other.d_time);
	//	System.out.println(" equal this f:"+this.fligth_num+" t:"+this.d_time);
		
		if(fligth_num.equals(other.fligth_num) && d_time.equals(other.d_time) &&
				this.departure.equals(other.departure) && this.arrival.equals(other.arrival))
		{  
	//		System.out.println("true");
			return true;
		}
		
	//	System.out.println("false");
		return false;
	}
	
	

}
