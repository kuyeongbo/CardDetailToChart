package card.codes;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.nio.charset.Charset;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

import com.opencsv.CSVReader;
import com.opencsv.exceptions.CsvException;

public class ReadCSV {

	ArrayList<HashMap<String, Object>> csvData = new ArrayList<HashMap<String, Object>>();

	private HashMap<String, Object> readCsvData(String path) throws IOException, CsvException {

		HashMap<String, Object> sheetData = new HashMap<String, Object>();
		CSVReader reader = new CSVReader(new FileReader(path, Charset.forName("SJIS")));
		File fileForName = new File(path);

		String year = fileForName.getName().substring(0, 4);
		String month = fileForName.getName().substring(4, 6);
		String sheetName = year + "年" + month + "月";

		ArrayList<CardDetailInfo> cardDetailInfoList = new ArrayList<CardDetailInfo>();
		ArrayList<String> catergoryList = new ArrayList<String>();
		ArrayList<String> sumPriceList = new ArrayList<String>();
		sheetData.put("sheetName", sheetName);

		List<String[]> ad = reader.readAll();
		for (String[] str : ad) {
			if (!str[0].equals("ＫＵ　ＹＥＯＮＧＢＯ　様")) {
				if (str[0].equals("") && str[1].equals("")) {
					CardDetailInfo cardInfo = new CardDetailInfo();
					cardInfo.setPrice(str[5]);
					cardDetailInfoList.add(cardInfo);
				} else if (!str[0].equals("") && !str[1].equals("") && !str[2].equals("")) {
					CardDetailInfo cardInfo = new CardDetailInfo();
					cardInfo.setDate(str[0]);
					cardInfo.setShop(str[1]);
					cardInfo.setPrice(str[2]);
					cardDetailInfoList.add(cardInfo);
				}
			}
		}
		String totalPrice = cardDetailInfoList.get(cardDetailInfoList.size() - 1).getPrice();
		sheetData.put("totalPrice", totalPrice);
		cardDetailInfoList = matchCatergory(cardDetailInfoList, sheetName);
		CardDetailInfo cdi = new CardDetailInfo();
		cdi.setDate("日 付");
		cdi.setShop("店 舗");
		cdi.setPrice("費 用");
		cardDetailInfoList.add(0, cdi);
		sheetData.put("cardDetailInfoList", cardDetailInfoList);
		for (int i = 0; i < cardDetailInfoList.size(); i++) {
			if (i != 0) {
				CardDetailInfo cardDetailInfo = new CardDetailInfo();
				cardDetailInfo = cardDetailInfoList.get(i);
				String categoryName = cardDetailInfo.getCategory();
				if (!(categoryName.equals("") || categoryName == null)) {
					int price = Integer.valueOf(cardDetailInfo.getPrice());
					if (!catergoryList.contains(categoryName)) {
						catergoryList.add(categoryName);
						sumPriceList.add(cardDetailInfo.getPrice());
					} else {
						int idx = catergoryList.indexOf(categoryName);
						int privPrice = Integer.valueOf(sumPriceList.get(idx));
						sumPriceList.set(idx, String.valueOf(privPrice + price));
					}
				}
			}
		}

		sheetData.put("catergoryList", catergoryList);
		sheetData.put("sumPriceList", sumPriceList);

		return sheetData;
	}

	public ArrayList<HashMap<String, Object>> getDataList() throws CsvException {

		ArrayList<String> filePathList = new ArrayList<String>();
		File folder = new File("C:\\Users\\kuyeo\\Desktop\\카드명세서\\カード明細書");
		filePathList = getFileList(folder, filePathList);
		HashMap<String, Object> map = new HashMap<String, Object>();
		csvData.add(0, map);
		csvData.get(0).put("nof", filePathList.size());

		for (String fileFullPath : filePathList) {
			try {
				HashMap<String, Object> sheetData = new HashMap<String, Object>();
				sheetData = readCsvData(fileFullPath);
				csvData.add(sheetData);
			} catch (IOException e) {
				// TODO 자동 생성된 catch 블록
				e.printStackTrace();
			}
		}
		return csvData;
	}

	public ArrayList<String> getFileList(File folder, ArrayList<String> filePathList) {
		for (File file : folder.listFiles()) {
			if (!file.isDirectory()) {
				filePathList.add(file.getAbsolutePath());
			} else {
				getFileList(file, filePathList);
			}
		}
		return filePathList;
	}

	public ArrayList<CardDetailInfo> matchCatergory(ArrayList<CardDetailInfo> CardDetailInfoList, String sheetName)
			throws IOException {

		File txtFile = new File("C:\\Users\\kuyeo\\Desktop\\카드명세서\\店舗リスト.txt");
		ArrayList<String> shopNameList = new ArrayList<String>();
		ArrayList<String> CategoryNameList = new ArrayList<String>();

		BufferedReader reader = new BufferedReader(new FileReader(txtFile));
		String shopNameInList;
		String catergoryNameInList;
		String line;
		while ((line = reader.readLine()) != null) {
			shopNameInList = line.split(",")[0];
			catergoryNameInList = line.split(",")[1];
			shopNameList.add(shopNameInList);
			CategoryNameList.add(catergoryNameInList);
		}
		reader.close();

		try {

			ArrayList<String> findCategoryNameList = new ArrayList<String>();
			ArrayList<String> unfindCategoryNameList = new ArrayList<String>();

			File unmatchedCategoryNameList = new File(
					"C:\\Users\\kuyeo\\Desktop\\카드명세서\\unmatchedCategoryNameList_" + sheetName + ".txt");

			for (int i = 0; i < CardDetailInfoList.size(); i++) {
				String shopName = CardDetailInfoList.get(i).getShop();
				if (!(shopName.equals("") || shopName == null)) {
					if (shopNameList.contains(shopName)) {
						int idx = shopNameList.indexOf(shopName);
						CardDetailInfoList.get(i).setCategory(CategoryNameList.get(idx));
						if (!findCategoryNameList.contains(shopName)) {
							findCategoryNameList.add(shopName);
						}
					} else {
						if (!unfindCategoryNameList.contains(shopName)) {
							unfindCategoryNameList.add(shopName);
						}
					}
				}
			}

			int unfindCategoryNameListNum = unfindCategoryNameList.size();
			if (unfindCategoryNameListNum > 0) {

				FileOutputStream fos = new FileOutputStream(unmatchedCategoryNameList, false);
				FileOutputStream fos2 = new FileOutputStream(txtFile, true);
				for (String name : unfindCategoryNameList) {
					String newCategory = name + "," + "カテゴリーを指定してください";
					fos2.write(newCategory.getBytes());
					fos2.write("\r\n".getBytes());
				}
				fos.close();
				fos2.close();
			}

		} finally {

		}
		return CardDetailInfoList;
	}
}
