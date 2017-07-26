package RecoveryPostSwap;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class SwapsPaser {

	public Map<String, String> subfleetMap = null;

	public List<SwapEntity> swapList = new ArrayList<SwapEntity>();;


	public  void paserData(String fileName){
		//System.out.println("Start...");
		BufferedReader br = null;
		StringBuffer buffer = null;
		
		String subFleetName = null;
		subfleetMap = new HashMap<String, String>();
		// read file
		try {
			File outFile = new File(fileName);
			buffer = new StringBuffer();
			InputStreamReader isr = new InputStreamReader(new FileInputStream(outFile), "utf-8");
			br = new BufferedReader(isr);
			String lineTxt = null;
			String lineRegEx = null;
			// Pattern linePattern null;// = Pattern.compile(lineRegEx);
			Matcher lineMatcher = null;

			while ((lineTxt = br.readLine()) != null) {
				//lineRegEx = "Subfleet ([0-9]{3}[a-zA-Z]*)";
				lineRegEx = "Subfleet ([0-9]{3})";
				Pattern linePattern = Pattern.compile(lineRegEx);
				lineMatcher = linePattern.matcher(lineTxt);

				if (lineMatcher.find()) {
					subFleetName = lineMatcher.group(1);
					//System.out.println(subFleetName +" -> "+lineMatcher.group(1));
				} else {
					lineRegEx = "Aircraft (B[0-9]{3,4})";
					Pattern acPattern = Pattern.compile(lineRegEx);
					lineMatcher = acPattern.matcher(lineTxt);
					if (lineMatcher.find()) {
						String acName = lineMatcher.group(1);
						subfleetMap.put(acName, subFleetName);
					}
				}

				buffer.append(lineTxt);
			}

		} catch (Exception e) {
			e.printStackTrace();
		}
		PaserSwaps(buffer.toString());
		//System.out.print(subfleetMap);

	}

	public List<SwapEntity> PaserSwaps(String inputTxt) {
		
		String regEx = "Aircraft (B[0-9]{3,4})[\\s\\S]{10,1000}?(Continues on its original path|switches to the original path of aircraft (B[0-9]{3,4}))";
		Pattern pattern = Pattern.compile(regEx);
		Matcher matcher = pattern.matcher(inputTxt);
		int pairCount = 0;
		while (matcher.find()) {
			pairCount++;
			if (!matcher.group(2).equals("Continues on its original path")) {
				swapList.add(new SwapEntity(matcher.group(3), matcher.group(1)));
			}
			// resultMap.get(matcher.group(1)).add(matcher.group(2));
			// System.out.println("group 1 " + matcher.group(1)
			// +matcher.group(2)+matcher.group(3));
			//System.out.println("group " + pairCount + matcher.group(1) + matcher.group(2) + matcher.group(3));
		}
		return swapList;
	}
}