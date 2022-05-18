import java.io.*;
import java.text.SimpleDateFormat;
import java.util.concurrent.TimeUnit;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
public class DatavyuToSupercoder {
	
	/**
	 * Open a given file and reads each entry one line at a time, creating Dayavyu objects with the attributes of 
	 * ordinal, onset, offset, and code. Insert each object into a list of Datavyu objects.
	 * @param fileName - the name of the csv file
	 * @return - a list of Datavyu objects which represent a single entry in the csv file
	 */
	public static List<DatavyuObject> readFile(String fileName) {
		List<DatavyuObject> lines = new LinkedList<>();
		
		try {
		    BufferedReader reader = new BufferedReader(new FileReader(fileName));
		    String line;
		    reader.readLine();
            while ((line = reader.readLine()) != null) {    
                String[] parts = line.split(","); // separate the line by comma
                int ordinal = Integer.parseInt(parts[0]); //convert string to integer
                int onset = Integer.parseInt(parts[1]); //convert string to integer
                int offset = Integer.parseInt(parts[2]);//convert string to integer
                String code = parts[3].replace("\"", "");//remove quotation from string
                DatavyuObject newData = new DatavyuObject(ordinal,onset,offset,code); //create object
                lines.add(newData);
            }            
            reader.close();   
		}
		catch (Exception e) {            
            System.err.format("Exception occurred trying to read '%s'.", fileName);            
            e.printStackTrace();        
        }
		return lines;
	}
	
	/**
	 * Given a list of data, list of start times, and file name, this function reformats the data 
	 * and writes it into a new csv file. The csv file will contain the updated onset/offset in frames
	 * and the start times of each trial given in frames, millisecond, and elapsed time. 
	 * @param dataList - a list of Datavyu objects which represents each entry in a csv file
	 * @param startTimes - a list of start times for each trial
	 * @param filename - name of the previous file that was read
	 */

	public static void writeFile(List<DatavyuObject> dataList, List<DatavyuObject> startTimes,String filename) {
		int count = 1;
		HSSFWorkbook workbook = new HSSFWorkbook();
		HSSFSheet sheet = workbook.createSheet("Sample sheet");

		Map <Integer, Object[] > data = new HashMap <Integer, Object[] > ();
		data.put(count, new Object[] {
				"Reformatted Data (in frames)",
				"",
				"",
				"",
				"",
				"Trial Start Times"
		});
		count++;
		data.put(count, new Object[] {
				"Code",
				"Onset",
				"Offset",
				"",
				"",
				"Trial Number",
				"Start Time (in frames - Supercoder)",
				"Start Time (in milliseconds - Datavyu CSVs)",
				"Start Time (in elapsed time - Datavyu coding)"
		});
		count++;
		int numTrials = startTimes.size();
	    int trialCount = 0; //keeps track of trial start times
	    boolean noOffset = false;
	    for(DatavyuObject d : dataList) { //iterating through all data in datalist
	    	String code = d.getCode(); 
		    double onset = d.getOnset();
		    double offset = d.getOffset();
		    int newOnset = (int)Math.round((onset/1000*29.97)); //create new onset in frames
		    int newOffset = 0;
		    if(offset == 0) {
		    	noOffset = true;
		    }
		    else {
			    newOffset = (int)Math.round(offset/1000*29.97); //create new offset in frames
			    noOffset = false;
			    
		    }
		    if(trialCount < numTrials) { //for inserting start times for each trial
		    	DatavyuObject cur = startTimes.get(trialCount);
		    	double onset2 = cur.getOnset();
		    	int newOnset2 = (int)Math.round((onset2/1000*29.97)); //create new start time in frames
		    	int millis = (int)Math.round(((double)newOnset2/29.97*1000)); //create new start time in milliseconds
			    	//sb.append(Integer.toString(newOnset2) +',');//add to csv file
			    	//sb.append(Integer.toString(millis) +',');//add to csv file
			    	
			    int minutes = millis /(1000 * 60);//get the number of minutes
			    int seconds = millis / 1000 % 60;//get the number of seconds
			    int milli  = millis % 1000;//get the number of milliseconds
			    String formatted = String.format("%02d:%02d.%03d", minutes, seconds, milli);//reformat as elapsed time	
			    //sb.append(formatted +',');//add elapsed time to csv file
			    if(noOffset) {
			    	data.put(count, new Object[] {
			    			code,
			    			newOnset,
			    			null,
			    			"",
			    			"",
			    			trialCount + 1,
			    			newOnset2,
			    			millis,
			    			formatted,
			    	});
			    	count++;
			    }
			    else {
			    	data.put(count, new Object[] {
			    			code,
			    			newOnset,
			    			newOffset,
			    			"",
			    			"",
			    			trialCount + 1,
			    			newOnset2,
			    			millis,
			    			formatted,
			    	});
			    	count++;
			    }
			    trialCount++;
			    continue;
		    		
		    }
		    if (noOffset && trialCount >= numTrials) {
		    	data.put(count, new Object[] {
		    			code,
		    			newOnset,
		    			null,
		    	});
		    	count++;
		    }
		    else if (!noOffset && trialCount >= numTrials) {
		    	data.put(count, new Object[]{
		    		code,
		    		newOnset,
		    		newOffset,
		    	});
		    	count++;
		    }
		    
		}
	

		Set <Integer> keyset = data.keySet();
		int rownum = 0;
		for (Integer key: keyset) {
		    Row row = sheet.createRow(rownum++);
		    Object[] objArr = data.get(key);
		    int cellnum = 0;
		    for (Object obj: objArr) {
		        Cell cell = row.createCell(cellnum++);
		        if(obj == null)
		        	continue;
		        if (obj instanceof Date)
		            cell.setCellValue((Date) obj);
		        else if (obj instanceof Boolean)
		            cell.setCellValue((Boolean) obj);
		        else if (obj instanceof String)
		            cell.setCellValue((String) obj);
		        else if (obj instanceof Double)
		            cell.setCellValue((Double) obj);
		        else if (obj instanceof Integer)
		        	cell.setCellValue(((Integer)obj));
		    }
		}

		try {
			filename = filename.substring(0,filename.length()-4);
		    FileOutputStream out =
		        new FileOutputStream(new File("Output/"+"OUTPUT_"+filename+".xlsx"));
		    workbook.write(out);
		    out.close();
		    System.out.println("Excel written successfully..");

		} catch (FileNotFoundException e) {
		    e.printStackTrace();
		} catch (IOException e) {
		    e.printStackTrace();
		}
	}
	/**
	 * Get the start time of each trial in the csv file
	 * @param dataList - list of Datavyu objects which represent each entry in the csv file
	 * @return - list of Datavyu objects that are the start time
	 */
	public static List<DatavyuObject> getStartTimes(List<DatavyuObject> dataList){
		List<DatavyuObject> startTimes = new LinkedList<>();
		for(DatavyuObject d : dataList) {
			String code = d.getCode();
			if(code.equals("B")) {
				startTimes.add(d);
			}
		}
		return startTimes;
		
	}
	/**
	 * Get all the file names from the Input folder
	 * @param folder - Input folder
	 * @return - a list of all file names within the Input folder
	 */
	public static List<String> getFiles(File folder){
		List<String> fileList = new ArrayList<>();
		for(File fileEntry : folder.listFiles()) {
			fileList.add(fileEntry.getName());
		}
		return fileList;
	}
	
	
	
	public static void main(String[] args) {
		File folder = new File("Input/"); //Input folder
        List<String> fileList = getFiles(folder); //get all file names within the Input folder
        for(String file : fileList) { //for each file name within the Input folder
        	String filePath = String.format("Input/%s",file); 
        	List<DatavyuObject> dataList = readFile(filePath);//read file
            List<DatavyuObject> startTimes = getStartTimes(dataList);//get start time
            writeFile(dataList,startTimes,file);//write new csv file with reformatted data
        }
	}

}
