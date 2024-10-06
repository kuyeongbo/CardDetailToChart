package card.codes;

import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;

import com.opencsv.exceptions.CsvException;

public class Main {

	public static void main(String[] args) throws InterruptedException, IOException {
		ArrayList<HashMap<String, Object>> csvDataList = new ArrayList<HashMap<String, Object>>();
		GEFBC GEFBC = new GEFBC();
		ReadCSV RCSV = new ReadCSV();
		try {
			csvDataList = RCSV.getDataList();
		} catch (CsvException e) {
			// TODO 자동 생성된 catch 블록
			e.printStackTrace();
		}
		GEFBC.getExcelFile(csvDataList);
		Thread.sleep(1000);
	}

}
