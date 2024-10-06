package card.codes;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jfree.data.category.DefaultCategoryDataset;

public class GetExcelFileByCSV {
	@SuppressWarnings({ "unchecked", "resource" })
	public void getExcelFile(ArrayList<HashMap<String, Object>> dataList) throws IOException {

		File xlsxFile = new File("C:\\Users\\kuyeo\\Desktop\\카드명세서\\test.xlsx");

		XSSFWorkbook workBook = new XSSFWorkbook();
		XSSFDataFormat format = workBook.createDataFormat();

		CellStyle percentStyle = workBook.createCellStyle();
		CellStyle defaultStyle = workBook.createCellStyle();
		CellStyle numberStyle = workBook.createCellStyle();

		percentStyle = setDefaultCellStyle(percentStyle);
		defaultStyle = setDefaultCellStyle(defaultStyle);
		numberStyle = setDefaultCellStyle(numberStyle);

		numberStyle.setDataFormat(format.getFormat("#,##0"));
		percentStyle.setDataFormat(format.getFormat("0%"));

//		int loopCnt = (int) dataList.get(0).get("nof");
		int loopCnt = 1;
		dataList.remove(0);

		for (int l = 0; l < loopCnt; l++) {

			ArrayList<CardDetailInfo> cardDetailInfoList = new ArrayList<CardDetailInfo>();
			ArrayList<String> catergoryList = (ArrayList<String>) dataList.get(l).get("catergoryList");
			ArrayList<String> sumPriceList = (ArrayList<String>) dataList.get(l).get("sumPriceList");
			cardDetailInfoList = (ArrayList<CardDetailInfo>) dataList.get(l).get("cardDetailInfoList");

			int totalPrice = Integer.valueOf(cardDetailInfoList.get(cardDetailInfoList.size() - 1).getPrice());

			// 시트 생성 및 셀 높이 설정
			String sheetName = (String) dataList.get(l).get("sheetName");
			XSSFSheet sheet = workBook.createSheet(sheetName);

			sheet.setDefaultRowHeightInPoints(30);

			for (int i = 0; i < cardDetailInfoList.size(); i++) {
				CardDetailInfo cardDetailInfo = new CardDetailInfo();
				cardDetailInfo = cardDetailInfoList.get(i);
				Row oriRow = sheet.createRow(i);
				for (int j = 0; j < dataList.size(); j++) {
					if (j == 0) {
						Cell dateCell = oriRow.createCell(j);
						dateCell.setCellStyle(defaultStyle);
						dateCell.setCellValue(cardDetailInfo.getDate());
						sheet.setColumnWidth(j, 3500);
					} else if (j == 1) {
						Cell shopCell = oriRow.createCell(j);
						shopCell.setCellStyle(defaultStyle);
						shopCell.setCellValue(cardDetailInfo.getShop());
						sheet.setColumnWidth(j, 3500);

					} else if (j == 2) {
						if (i > 0) {
							Cell priceCell = oriRow.createCell(j);
							priceCell.setCellStyle(numberStyle);
							priceCell.setCellValue(cardDetailInfo.getPrice());
							sheet.setColumnWidth(j, 3500);
						} else {
							Cell priceCell = oriRow.createCell(j);
							priceCell.setCellStyle(defaultStyle);
							priceCell.setCellValue(cardDetailInfo.getPrice());
							sheet.setColumnWidth(j, 3500);
						}
					}
				}
			}

			DefaultCategoryDataset cardData = new DefaultCategoryDataset();

			for (int i = 0; i < catergoryList.size(); i++) {
				Row oriRow = sheet.getRow(i);
				int cells = oriRow.getPhysicalNumberOfCells();
				for (int j = 0; j < 3; j++) {
					float percentage = (float) Integer.valueOf(sumPriceList.get(i)) / (float) totalPrice;
					double doublePercentage = (double) Integer.valueOf(sumPriceList.get(i)) / (double) totalPrice;
					int sumPrice = Integer.valueOf(sumPriceList.get(i));
					String catergory = catergoryList.get(i);
					cardData.addValue(doublePercentage, "Chart", catergory);
					if (j == 0) {
						Cell catergoryCell = oriRow.createCell(cells + 1);
						catergoryCell.setCellStyle(defaultStyle);
						catergoryCell.setCellValue(catergory);
						sheet.setColumnWidth(j, 3500);
					} else if (j == 1) {
						Cell sumPriceCell = oriRow.createCell(cells + 2);
						sumPriceCell.setCellStyle(numberStyle);
						sumPriceCell.setCellValue(sumPrice);
						sheet.setColumnWidth(j, 3500);
					} else if (j == 2) {
						Cell perCell = oriRow.createCell(cells + 3);
						perCell.setCellStyle(percentStyle);
//						float percentage = (float) Integer.valueOf(sumPriceList.get(i)) / (float) totalPrice;
						perCell.setCellValue(percentage);
						sheet.setColumnWidth(j, 3500);
					}
				}
			}

			GetChart chartGenerator = new GetChart();
//			GetChartByJFreeChart GCBJC = new GetChartByJFreeChart();

			workBook = chartGenerator.generateBarChart(workBook, sheetName, catergoryList.size() - 1);

		}

		FileOutputStream fileOut = new FileOutputStream(xlsxFile);
		try {
			workBook.write(fileOut);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			fileOut.close();
			workBook.close();
			System.out.println("save_Complete");
		}
	}

	private CellStyle setDefaultCellStyle(CellStyle cellStyle) {

		// 테두리 설정
		cellStyle.setBorderTop(BorderStyle.THIN);
		cellStyle.setBorderLeft(BorderStyle.THIN);
		cellStyle.setBorderRight(BorderStyle.THIN);
		cellStyle.setBorderBottom(BorderStyle.THIN);

		// 줄 바꿈 및 중앙 정렬
		cellStyle.setWrapText(true);
		cellStyle.setAlignment(HorizontalAlignment.CENTER);
		cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

		return cellStyle;
	}
}
