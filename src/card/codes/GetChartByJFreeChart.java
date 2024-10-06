package card.codes;

import java.awt.Font;
import java.io.ByteArrayOutputStream;
import java.text.DecimalFormat;

import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartUtilities;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.StandardChartTheme;
import org.jfree.chart.axis.CategoryAxis;
import org.jfree.chart.axis.CategoryLabelPositions;
import org.jfree.chart.axis.NumberAxis;
import org.jfree.chart.labels.StandardCategoryItemLabelGenerator;
import org.jfree.chart.plot.CategoryPlot;
import org.jfree.chart.plot.PlotOrientation;
import org.jfree.chart.renderer.category.BarRenderer;
import org.jfree.data.category.DefaultCategoryDataset;

public class GetChartByJFreeChart {

	public XSSFWorkbook generateBarChart(XSSFWorkbook workBook, String sheetName, DefaultCategoryDataset cardData) {

		try {
			XSSFSheet sheet = workBook.getSheet(sheetName);
			XSSFDrawing drawing = sheet.createDrawingPatriarch();

			ChartFactory.setChartTheme(StandardChartTheme.createLegacyTheme());
//			JFreeChart barChartObject = ChartFactory.createBarChart("カード支出内訳", "カテゴリー", "比率", cardData,
//					PlotOrientation.VERTICAL, true, true, false);
			JFreeChart barChartObject = ChartFactory.createBarChart("", "", "", cardData, PlotOrientation.VERTICAL,
					true, true, false);

			int numOfCardData = cardData.getColumnCount();

			int width = numOfCardData * 600;
			int fontSize = width / 100;
			int height = 768;

			// Bar Renderer 설정 (퍼센트 표시)
			CategoryPlot plot = barChartObject.getCategoryPlot();

			NumberAxis yAxis = (NumberAxis) plot.getRangeAxis();
			CategoryAxis xAxis = plot.getDomainAxis();
			BarRenderer renderer = (BarRenderer) plot.getRenderer();

//			xAxis.setTickLabelFont(new Font("SansSerif", Font.PLAIN, 30)); // 여기서 12는 폰트 크기입니다
//			yAxis.setTickLabelFont(new Font("SansSerif", Font.PLAIN, 30));
//			renderer.setBaseItemLabelFont(new Font("SansSerif", Font.PLAIN, 30)); // Bar Renderer 폰트 크기 설정

			xAxis.setTickLabelFont(new Font("SansSerif", Font.PLAIN, fontSize)); // 여기서 12는 폰트 크기입니다
			yAxis.setTickLabelFont(new Font("SansSerif", Font.PLAIN, fontSize));
			renderer.setBaseItemLabelFont(new Font("SansSerif", Font.PLAIN, fontSize)); // Bar Renderer 폰트 크기 설정

			yAxis.setNumberFormatOverride(new DecimalFormat("0.0%"));

			renderer.setBaseItemLabelGenerator(
					new StandardCategoryItemLabelGenerator("{2}", new DecimalFormat("0.0%")));
			renderer.setBaseItemLabelsVisible(true);

			ByteArrayOutputStream baos = new ByteArrayOutputStream();
			ChartUtilities.writeChartAsPNG(baos, barChartObject, width, height);

			int my_pic_id = workBook.addPicture(baos.toByteArray(), Workbook.PICTURE_TYPE_PNG);
			baos.close();

			int rowStartPoint = 8;
			int columnStartPoint = 0;
			int rowEndPoint = rowStartPoint + (numOfCardData * 3);
			int columnEndPoint = 15;

			ClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, rowStartPoint, columnStartPoint, rowEndPoint,
					columnEndPoint);
			drawing.createPicture(anchor, my_pic_id);

		} catch (Exception e) {
			e.printStackTrace();
			System.out.println("차트 생성 중 오류 발생: " + e.getMessage());
		}
		return workBook;
	}
}
