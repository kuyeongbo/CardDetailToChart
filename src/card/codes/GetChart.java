package card.codes;

import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.AxisCrosses;
import org.apache.poi.xddf.usermodel.chart.AxisPosition;
import org.apache.poi.xddf.usermodel.chart.ChartTypes;
import org.apache.poi.xddf.usermodel.chart.XDDFCategoryAxis;
import org.apache.poi.xddf.usermodel.chart.XDDFChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFChartLegend;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory;
import org.apache.poi.xddf.usermodel.chart.XDDFNumericalDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFValueAxis;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTAxDataSource;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBarChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBarSer;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBoolean;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTCatAx;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTLegend;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTNumDataSource;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTNumRef;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPlotArea;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTScaling;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTSerTx;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTStrRef;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTValAx;
import org.openxmlformats.schemas.drawingml.x2006.chart.STAxPos;
import org.openxmlformats.schemas.drawingml.x2006.chart.STBarDir;
import org.openxmlformats.schemas.drawingml.x2006.chart.STLegendPos;
import org.openxmlformats.schemas.drawingml.x2006.chart.STOrientation;
import org.openxmlformats.schemas.drawingml.x2006.chart.STTickLblPos;

public class GetChart {

//	private XSSFWorkbook workBook;
//	private XSSFSheet sheet;
//	private int size;

//	public GetChart() {
//		this.workBook = workBook;
//		this.sheet = sheet;
//		this.size = size;
//	}

	public XSSFWorkbook generateBarChart(XSSFWorkbook workBook, String sheetName, int size) {

		XSSFSheet sheet = workBook.getSheet(sheetName);
		// Create a drawing patriarch to anchor the chart in the worksheet
		XSSFDrawing drawing = sheet.createDrawingPatriarch();

		// Create an anchor with upper-left cell and lower-right cell
		ClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 8, 0, 22, 15);

		// Create a chart (casting the drawing to XSSFDrawing)
		try {
			XSSFChart chart = drawing.createChart(anchor);

			CTChart ctChart = ((XSSFChart) chart).getCTChart();
			CTPlotArea ctPlotArea = ctChart.getPlotArea();
			CTBarChart ctBarChart = ctPlotArea.addNewBarChart();
			CTBoolean ctBoolean = ctBarChart.addNewVaryColors();
			ctBoolean.setVal(true);
			ctBarChart.addNewBarDir().setVal(STBarDir.COL);

//			for (int r = 2; r < 6; r++) {
			CTBarSer ctBarSer = ctBarChart.addNewSer();
			CTSerTx ctSerTx = ctBarSer.addNewTx();
			CTStrRef ctStrRef = ctSerTx.addNewStrRef();
//			ctStrRef.setF("Sheet1!$A$1");
			ctStrRef.setF("Sheet1!$D$1");
			ctBarSer.addNewIdx().setVal(0);
			CTAxDataSource cttAxDataSource = ctBarSer.addNewCat();
			ctStrRef = cttAxDataSource.addNewStrRef();
//			ctStrRef.setF("Sheet1!$B$1:$G$1");
			ctStrRef.setF("Sheet1!$E$1:$E$"+size);
			CTNumDataSource ctNumDataSource = ctBarSer.addNewVal();
			CTNumRef ctNumRef = ctNumDataSource.addNewNumRef();
//			ctNumRef.setF("Sheet1!$B$2:$G$2");
			ctNumRef.setF("Sheet1!$G$1:$G$"+size);

				// at least the border lines in Libreoffice Calc ;-)
				ctBarSer.addNewSpPr().addNewLn().addNewSolidFill().addNewSrgbClr().setVal(new byte[] { 0, 0, 0 });

//			}

			// telling the BarChart that it has axes and giving them Ids
			ctBarChart.addNewAxId().setVal(123456);
			ctBarChart.addNewAxId().setVal(123457);

			// cat axis
			CTCatAx ctCatAx = ctPlotArea.addNewCatAx();
			ctCatAx.addNewAxId().setVal(123456); // id of the cat axis
			CTScaling ctScaling = ctCatAx.addNewScaling();
			ctScaling.addNewOrientation().setVal(STOrientation.MIN_MAX);
			ctCatAx.addNewDelete().setVal(false);
			ctCatAx.addNewAxPos().setVal(STAxPos.B);
			ctCatAx.addNewCrossAx().setVal(123457); // id of the val axis
			ctCatAx.addNewTickLblPos().setVal(STTickLblPos.NEXT_TO);

			// val axis
			CTValAx ctValAx = ctPlotArea.addNewValAx();
			ctValAx.addNewAxId().setVal(123457); // id of the val axis
			ctScaling = ctValAx.addNewScaling();
			ctScaling.addNewOrientation().setVal(STOrientation.MIN_MAX);
			ctValAx.addNewDelete().setVal(false);
			ctValAx.addNewAxPos().setVal(STAxPos.L);
			ctValAx.addNewCrossAx().setVal(123456); // id of the cat axis
			ctValAx.addNewTickLblPos().setVal(STTickLblPos.NEXT_TO);

			// legend
			CTLegend ctLegend = ctChart.addNewLegend();
			ctLegend.addNewLegendPos().setVal(STLegendPos.B);
			ctLegend.addNewOverlay().setVal(false);

			System.out.println(ctChart);
			
		} catch (Exception e) {
			e.printStackTrace();
			System.out.println("차트 생성 중 오류 발생: " + e.getMessage());
		}
		return workBook;
	}
}
