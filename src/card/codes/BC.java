package card.codes;

import java.io.FileOutputStream;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDrawing;

import org.openxmlformats.schemas.drawingml.x2006.chart.CTChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPlotArea;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBarChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBoolean;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBarSer;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTAxDataSource;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTNumDataSource;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTNumRef;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTStrRef;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTSerTx;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTCatAx;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTValAx;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTScaling;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTLegend;
import org.openxmlformats.schemas.drawingml.x2006.chart.STAxPos;
import org.openxmlformats.schemas.drawingml.x2006.chart.STBarDir;
import org.openxmlformats.schemas.drawingml.x2006.chart.STCrosses;
import org.openxmlformats.schemas.drawingml.x2006.chart.STOrientation;
import org.openxmlformats.schemas.drawingml.x2006.chart.STLegendPos;
import org.openxmlformats.schemas.drawingml.x2006.chart.STTickLblPos;

public class BC {

	public static void main(String[] args) throws Exception {
		Workbook wb = new XSSFWorkbook("C:\\Users\\kuyeo\\Desktop\\카드명세서\\test.xlsx");
//		Workbook wb = new XSSFWorkbook("C:\\Users\\kuyeo\\Desktop\\카드명세서\\BarChart_2.xlsx");
//		Workbook wb = new XSSFWorkbook("C:\\Users\\kuyeo\\Desktop\\카드명세서\\BarChart_3.xlsx");
		Sheet sheet = wb.getSheetAt(0);

		XSSFDrawing drawing = (XSSFDrawing) sheet.createDrawingPatriarch();
		ClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 8, 0, 22, 15);

		XSSFChart chart = drawing.createChart(anchor);

		CTChart ctChart = ((XSSFChart) chart).getCTChart();
		CTPlotArea ctPlotArea = ctChart.getPlotArea();
		CTBarChart ctBarChart = ctPlotArea.addNewBarChart();
		CTBoolean ctBoolean = ctBarChart.addNewVaryColors();
		ctBoolean.setVal(true);
		ctBarChart.addNewBarDir().setVal(STBarDir.COL);

//		for (int r = 2; r < 6; r++) {
//			CTBarSer ctBarSer = ctBarChart.addNewSer();
//			CTSerTx ctSerTx = ctBarSer.addNewTx();
//			CTStrRef ctStrRef = ctSerTx.addNewStrRef();
//			ctStrRef.setF("Sheet1!$A$" + r);
//			ctBarSer.addNewIdx().setVal(r - 2);
//			CTAxDataSource cttAxDataSource = ctBarSer.addNewCat();
//			ctStrRef = cttAxDataSource.addNewStrRef();
//			ctStrRef.setF("Sheet1!$B$1:$D$1");
//			CTNumDataSource ctNumDataSource = ctBarSer.addNewVal();
//			CTNumRef ctNumRef = ctNumDataSource.addNewNumRef();
//			ctNumRef.setF("Sheet1!$B$" + r + ":$D$" + r);
//
//			// at least the border lines in Libreoffice Calc ;-)
//			ctBarSer.addNewSpPr().addNewLn().addNewSolidFill().addNewSrgbClr().setVal(new byte[] { 0, 0, 0 });
//
//		}
		
//		for (int r = 2; r < 6; r++) {
//			CTBarSer ctBarSer = ctBarChart.addNewSer();
//			CTSerTx ctSerTx = ctBarSer.addNewTx();
//			CTStrRef ctStrRef = ctSerTx.addNewStrRef();
//			ctStrRef.setF("Sheet1!$A$" + r);
//			ctBarSer.addNewIdx().setVal(r - 2);
//			CTAxDataSource cttAxDataSource = ctBarSer.addNewCat();
//			ctStrRef = cttAxDataSource.addNewStrRef();
//			ctStrRef.setF("Sheet1!$B$1:$D$1");
//			CTNumDataSource ctNumDataSource = ctBarSer.addNewVal();
//			CTNumRef ctNumRef = ctNumDataSource.addNewNumRef();
//			ctNumRef.setF("Sheet1!$B$" + r + ":$D$" + r);
//
//			// at least the border lines in Libreoffice Calc ;-)
//			ctBarSer.addNewSpPr().addNewLn().addNewSolidFill().addNewSrgbClr().setVal(new byte[] { 0, 0, 0 });
//
//		}
		
//		for (int r = 2; r < 7; r++) {
			CTBarSer ctBarSer = ctBarChart.addNewSer();
			CTSerTx ctSerTx = ctBarSer.addNewTx();
			CTStrRef ctStrRef = ctSerTx.addNewStrRef();
//			ctStrRef.setF("Sheet1!$A$1");
			ctStrRef.setF("Sheet1!$D$1");
			ctBarSer.addNewIdx().setVal(0);
			CTAxDataSource cttAxDataSource = ctBarSer.addNewCat();
			ctStrRef = cttAxDataSource.addNewStrRef();
//			ctStrRef.setF("Sheet1!$B$1:$G$1");
			ctStrRef.setF("Sheet1!$E$2:$E$7");
			CTNumDataSource ctNumDataSource = ctBarSer.addNewVal();
			CTNumRef ctNumRef = ctNumDataSource.addNewNumRef();
//			ctNumRef.setF("Sheet1!$B$2:$G$2");
			ctNumRef.setF("Sheet1!$G$2:$G$7");

			// at least the border lines in Libreoffice Calc ;-)
			ctBarSer.addNewSpPr().addNewLn().addNewSolidFill().addNewSrgbClr().setVal(new byte[] { 0, 0, 0 });

//		}

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

		FileOutputStream fileOut = new FileOutputStream("C:\\Users\\kuyeo\\Desktop\\카드명세서\\BarChart.xlsx");
		wb.write(fileOut);
		fileOut.close();
	}
}