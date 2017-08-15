import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.Date;

import sun.util.logging.resources.logging;
/**
 * 
 * @author LancelotWG
 *
 */
enum BrokenStatus{
	cancel,delay,swap,normal
}
public class Flights implements Comparable<Flights> {
	Number number;
	ArrayList<Place> places = new ArrayList<>(); 
	boolean isBreak = false;
	public Flights(Number number, Place place) {
		this.number = number;
		places.add(place);
	}
	@Override
	public String toString() {	
		return "Flight number :"+this.number.flightNumber+"tailNumber :"+places.get(0).tailNumber+
				",Date :"+places.get(0).departureTime.toString()+ "-" + places.get(places.size() - 1).arrivalTime.toString()
				+ ",From :"+places.get(0).departure+",To :"+places.get(places.size() - 1).arrival;
	}
	@Override
	public int compareTo(Flights arg0) {
		Flights other = arg0;
		if(number.flightNumber.equals(other.number.flightNumber)){
			if(places.size() <= 1){
				if(places.get(0).arrival.equals(other.places.get(0).departure) ||
						places.get(0).departure.equals(other.places.get(0).arrival)){
					if(places.get(0).arrival.equals(other.places.get(0).departure)){
						if(other.places.get(0).departureTime.getTime() - places.get(0).arrivalTime.getTime() > 0){
							if(other.places.get(0).departureTime.getYear() == places.get(0).arrivalTime.getYear() &&
									other.places.get(0).departureTime.getMonth() == places.get(0).arrivalTime.getMonth() &&
									other.places.get(0).departureTime.getDay() == places.get(0).arrivalTime.getDay()){
								other.places.addAll(places);
								java.util.Collections.sort(other.places,new SortByTime());
								return 0;
							}else{
								return -1;
							}
						}else{
							return -1;
						}
					}else{
						if(places.get(0).departureTime.getTime() - other.places.get(0).arrivalTime.getTime() > 0){
							if(places.get(0).departureTime.getYear() == other.places.get(0).arrivalTime.getYear() &&
									places.get(0).departureTime.getMonth() == other.places.get(0).arrivalTime.getMonth() &&
									places.get(0).departureTime.getDay() == other.places.get(0).arrivalTime.getDay()){
								other.places.addAll(places);
								java.util.Collections.sort(other.places,new SortByTime());
								return 0;
							}else{
								return -1;
							}
						}else{
							return -1;
						}
					}

				}else{
					return -1;
				}
			}else{
				if(places.get(places.size() - 1).arrival.equals(other.places.get(0).departure)){
					if(places.get(places.size() - 1).arrivalTime.before(other.places.get(0).departureTime)){
						if(other.places.get(0).departureTime.getYear() == places.get(places.size() - 1).arrivalTime.getYear() &&
								other.places.get(0).departureTime.getMonth() == places.get(places.size() - 1).arrivalTime.getMonth() &&
								other.places.get(0).departureTime.getDay() == places.get(places.size() - 1).arrivalTime.getDay()){
							other.places.addAll(places);
							java.util.Collections.sort(other.places,new SortByTime());
							return 0;
						}else{
							return -1;
						}
					}else{
						return -1;
					}
				}else{
					return -1;
				}
			}
		}else{
			return -1;
		}
	}
	@Override
	public int hashCode() {
		return this.number.flightNumber.hashCode();
	}
	@Override
	public boolean equals(Object arg0) {
		Flights other=(Flights)arg0;
		if(number.flightNumber.equals(other.number.flightNumber)){
			if(places.size() <= 1){
				if(places.get(0).arrival.equals(other.places.get(0).departure) ||
						places.get(0).departure.equals(other.places.get(0).arrival)){
					if(places.get(0).arrival.equals(other.places.get(0).departure)){
						if(other.places.get(0).departureTime.getTime() - places.get(0).arrivalTime.getTime() > 0){
							if(other.places.get(0).departureTime.getYear() == places.get(0).arrivalTime.getYear() &&
									other.places.get(0).departureTime.getMonth() == places.get(0).arrivalTime.getMonth() &&
									other.places.get(0).departureTime.getDay() == places.get(0).arrivalTime.getDay()){
								other.places.addAll(places);
								java.util.Collections.sort(other.places,new SortByTime());
								return true;
							}else{
								return false;
							}
						}else{
							return false;
						}
					}else{
						if(places.get(0).departureTime.getTime() - other.places.get(0).arrivalTime.getTime() > 0){
							if(places.get(0).departureTime.getYear() == other.places.get(0).arrivalTime.getYear() &&
									places.get(0).departureTime.getMonth() == other.places.get(0).arrivalTime.getMonth() &&
									places.get(0).departureTime.getDay() == other.places.get(0).arrivalTime.getDay()){
								other.places.addAll(places);
								java.util.Collections.sort(other.places,new SortByTime());
								return true;
							}else{
								return false;
							}
						}else{
							return false;
						}
					}

				}else{
					return false;
				}
			}else{
				if(places.get(places.size() - 1).arrival.equals(other.places.get(0).departure)){
					if(places.get(places.size() - 1).arrivalTime.before(other.places.get(0).departureTime)){
						if(other.places.get(0).departureTime.getYear() == places.get(places.size() - 1).arrivalTime.getYear() &&
								other.places.get(0).departureTime.getMonth() == places.get(places.size() - 1).arrivalTime.getMonth() &&
								other.places.get(0).departureTime.getDay() == places.get(places.size() - 1).arrivalTime.getDay()){
							other.places.addAll(places);
							java.util.Collections.sort(other.places,new SortByTime());
							return true;
						}else{
							return false;
						}
					}else{
						return false;
					}
				}else{
					return false;
				}
			}
		}else{
			return false;
		}
		/*if(number.flightNumber.equals(other.number.flightNumber) && number.tailNumber.equals(other.number.tailNumber)) 		
		{
			if(places.size() <= 1){
				if(places.get(0).arrival.equals(other.places.get(0).departure) ||
						places.get(0).departure.equals(other.places.get(0).arrival)){
					if(places.get(0).arrival.equals(other.places.get(0).departure)){
						if(other.places.get(0).departureTime.getTime() - places.get(0).arrivalTime.getTime() > 0){
							long timeout = 2*60*60*1000;
							long intime = 20*60*1000;
							if(other.places.get(0).departureTime.getTime() - places.get(0).arrivalTime.getTime() <= timeout &&
									other.places.get(0).departureTime.getTime() - places.get(0).arrivalTime.getTime() >= intime){
								other.places.addAll(places);
								java.util.Collections.sort(other.places,new SortByTime());
								return true;
							}else{
								return false;
							}
						}else{
							return false;
						}
					}else{
						if(places.get(0).departureTime.getTime() - other.places.get(0).arrivalTime.getTime() > 0){
							long timeout = 2*60*60*1000;
							long intime = 20*60*1000;
							if(places.get(0).departureTime.getTime() - other.places.get(0).arrivalTime.getTime() <= timeout &&
									places.get(0).departureTime.getTime() - other.places.get(0).arrivalTime.getTime() >= intime){
								other.places.addAll(places);
								java.util.Collections.sort(other.places,new SortByTime());
								return true;
							}else{
								return false;
							}
						}else{
							return false;
						}
					}

				}else{
					return false;
				}
			}else{
				if(places.get(places.size() - 1).arrival.equals(other.places.get(0).departure)){
					if(places.get(places.size() - 1).arrivalTime.before(other.places.get(0).departureTime)){
						long timeout = 2*60*60*1000;
						long intime = 20*60*1000;
						if(other.places.get(0).departureTime.getTime() - places.get(places.size() - 1).arrivalTime.getTime() <= timeout &&
								other.places.get(0).departureTime.getTime() - places.get(places.size() - 1).arrivalTime.getTime() >= intime){
							other.places.addAll(places);
							java.util.Collections.sort(other.places,new SortByTime());
							return true;
						}else{
							return false;
						}
					}else{
						return false;
					}
				}else{
					return false;
				}
			}
			
		}else{
			return false;
		}*/
	}
}
class BrokenRecords{
	Number number;
	ArrayList<Place> places = new ArrayList<>(); 
	public BrokenRecords(Number number, ArrayList<Place> places){
		this.number = number;
		this.places = places;
	}
	public Number getNumber() {
		return number;
	}
}
class Number{
	String flightNumber;
	public Number(String flightNumber){
		this.flightNumber = flightNumber;
	}
	public String getFlightNumber() {
		return flightNumber;
	}
	public void setFlightNumber(String flightNumber) {
		this.flightNumber = flightNumber;
	}
}
class Place implements Comparable<Place>{
	String tailNumber;
	String arrival;
	String departure;
	Date arrivalTime;
	Date departureTime;
	BrokenStatus brokenStatus = BrokenStatus.normal;
	String newFlight = null;
	Date expectedDepartureTime = null;
	Date expectedArrivalTime = null;
	String pax = null;
	public Place(String tailNumber, String arrival, String departure, String arrivalTime, String departureTime){
		this.tailNumber = tailNumber;
		this.arrival = arrival;
		this.departure = departure;
		SimpleDateFormat format = new SimpleDateFormat("ddMMyyyy_HHmm");
		Date arrivalDate = null;
		try {
			arrivalDate = format.parse(arrivalTime);
		} catch (ParseException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}
		Date departureDate = null;
		try {
			departureDate = format.parse(departureTime);
		} catch (ParseException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		this.arrivalTime = arrivalDate;
		this.departureTime = departureDate;
	}
	public String getArrival() {
		return arrival;
	}
	public void setArrival(String arrival) {
		this.arrival = arrival;
	}
	public String getDeparture() {
		return departure;
	}
	public void setDeparture(String departure) {
		this.departure = departure;
	}
	public Date getArrivalTime() {
		return arrivalTime;
	}
	public void setArrivalTime(Date arrivalTime) {
		this.arrivalTime = arrivalTime;
	}
	public Date getDepartureTime() {
		return departureTime;
	}
	public void setDepartureTime(Date departureTime) {
		this.departureTime = departureTime;
	}
	@Override
	public int compareTo(Place arg0) {
		if(arrivalTime.before(arg0.departureTime)) 		
		{//   System.out.println("true");
			return -1;
		}else{
			return 1;
		}
	}
	public BrokenStatus getBrokenStatus() {
		return brokenStatus;
	}
	public void setBrokenStatus(BrokenStatus brokenStatus) {
		this.brokenStatus = brokenStatus;
	}
	public String getTailNumber() {
		return tailNumber;
	}
	public void setTailNumber(String tailNumber) {
		this.tailNumber = tailNumber;
	}
	public String getNewFlight() {
		return newFlight;
	}
	public void setNewFlight(String newFlight) {
		this.newFlight = newFlight;
	}
	public Date getExpectedDepartureTime() {
		return expectedDepartureTime;
	}
	public void setExpectedDepartureTime(Date expectedDepartureTime) {
		this.expectedDepartureTime = expectedDepartureTime;
	}
	public Date getExpectedArrivalTime() {
		return expectedArrivalTime;
	}
	public void setExpectedArrivalTime(Date expectedArrivalTime) {
		this.expectedArrivalTime = expectedArrivalTime;
	}
	public String getPax() {
		return pax;
	}
	public void setPax(String pax) {
		this.pax = pax;
	}
}
class SortByTime implements Comparator<Place>{
	@Override
	public int compare(Place arg0, Place arg1) {
		// TODO Auto-generated method stub
		if(arg1.departureTime.getTime() - arg0.arrivalTime.getTime() > 0){
			return -1;
		}else{
			return 1;
		}
	}
}