package com.euscold.testcases;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Method;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.concurrent.TimeUnit;

import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.Select;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import com.euscold.base.DBConnection;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

public class Inventory_Inquiry_Multiple_Prods extends DBConnection{
		
	@BeforeTest
	public void Set_up() {
		try {
			String dbClass = "com.ibm.db2.jcc.DB2Driver";
			Class.forName(dbClass).newInstance();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	@Test
	public void excelCreation() throws IOException {
		fileout = new FileOutputStream(test_results+"Inventory_Inquiry_Prod.xls");
		sheet = WB.createSheet("Product");
	}

	public void PRD(int row,String WHSE, String CUST, String Prd_Code) throws SQLException, IOException {
		rowhead = sheet.createRow(0);
		rowhead.createCell(0).setCellValue("Product Code");
		rowhead.createCell(1).setCellValue("Product Desc");
		rowhead.createCell(2).setCellValue("Reserved_Qty");
		rowhead.createCell(3).setCellValue("PIP Qty");
		rowhead.createCell(4).setCellValue("Avail Qty");
		rowhead.createCell(5).setCellValue("Hold Qty");
		rowhead.createCell(6).setCellValue("Total Qty Ord");
		rowhead.createCell(7).setCellValue("Total_Qty_Avail_Ord");
		rowhead.createCell(8).setCellValue("Potential_Shot_Qty");
		rowhead.createCell(9).setCellValue("Rec_In_Proc_Qty");
		rowhead.createCell(10).setCellValue("Exp_Inb_Qty");
		rowhead.createCell(11).setCellValue("Total Qty");
		datasource.setUser("HQIRCHEN");
		datasource.setPassword("HQIRCHEN");
		Connection connection = datasource.getConnection();
		stmt = connection.createStatement();
		String query = "SELECT PROD_CODE,PROD_DESC FROM MIGSRCMIL.V_WHSE_PROD as VWP " + "WHERE VWP.WHSE_SYSID=" + WHSE
				+ " AND VWP.CUST_SYSID=" + CUST + " AND VWP.PROD_CODE='" + Prd_Code + "' ";
		System.out.println(query); // 98403 98596
		ResultSet res = stmt.executeQuery(query);
		System.out.println("\n");

		if (res.next()) {
			rowhead = sheet.createRow((short) row);
			rowhead.createCell(0).setCellValue(res.getString(("PROD_CODE")));
			rowhead.createCell(1).setCellValue(res.getString("PROD_DESC"));
			System.out.println(String.format("%s - %s - ", res.getString(1), res.getString(2)));
		}else {
			rowhead = sheet.createRow((short) row);
		}
		System.out.println("Porduct code and Product Description:");

	}

	public void Reserved_Qty(String WHSE, String CUST, String Product_Code) throws SQLException, IOException {

		
		String query = "SELECT TRIM(CASE WHEN to_char((A.TASK_Reserved_Qty+B.ORDERED_QTY+C.C_ORDERED_QTY),'9,999') is null THEN '0' "
				+ "ELSE to_char((A.TASK_Reserved_Qty+B.ORDERED_QTY+C.C_ORDERED_QTY),'9,999') END) Reserved_Qty FROM "
				+ "(SELECT coalesce(SUM(COALESCE(PTD.QTY_TO_PICK,0)),0) AS TASK_Reserved_Qty "
				+ "FROM   MIGSRCMIL.ORDER as ORDR "
				+ "Inner Join MIGSRCMIL.ORDER_DTL AS ORD_DTL ON ORDR.DLVRY_TICKET_SYSID = ORD_DTL.DLVRY_TICKET_SYSID "
				+ "Inner Join MIGSRCMIL.PICK_TASK_DTL as PTD ON PTD.DLVRY_TICKET_SYSID = ORD_DTL.DLVRY_TICKET_SYSID  AND "
				+ "PTD.ORDER_DTL_SYSID =ORD_DTL.ORDER_DTL_SYSID "
				+ "LEFT JOIN MIGSRCMIL.V_LOT_HEADER as LH  ON LH.USCS_LOT_NUMBER = ORD_DTL.USCS_LOT_NUMBER "
				+ "LEFT JOIN MIGSRCMIL.DTL_CODE AS DT_MDE ON ORD_DTL.METHOD_OF_DTL_ENTRY_SYSID=DT_MDE.DTL_CODE_SYSID "
				+ "LEFT JOIN MIGSRCMIL.DTL_CODE AS DT_ST ON ORDR.ORDER_STATUS_SYSID=DT_ST.DTL_CODE_SYSID "
				+ "INNER JOIN MIGSRCMIL.TASK_HEADER as TH ON TH.TASK_NUMBER_SYSID = PTD.TASK_NUMBER_SYSID "
				+ "INNER JOIN MIGSRCMIL.DTL_CODE as DTL_TSK_TYP ON TH.TASK_TYPE_SYSID = DTL_TSK_TYP.DTL_CODE_SYSID AND DTL_TSK_TYP.DTL_CODE='BK' "
				+ "INNER JOIN MIGSRCMIL.V_WHSE_PROD as VWP ON VWP.PROD_SYSID=ORD_DTL.PROD_SYSID AND VWP.WHSE_SYSID=ORD_DTL.WHSE_SYSID AND  "
				+ "VWP.CUST_SYSID=ORD_DTL.CUST_SYSID "
				+ " WHERE  DT_MDE.DTL_CODE IN ('PRD','CDT','PSQ')  AND PTD.PROJECTION_FLAG = 'N' AND DT_ST.DTL_CODE='0' and "
				+ "VWP.WHSE_SYSID=" + WHSE + " AND VWP.CUST_SYSID=" + CUST + " AND VWP.PROD_CODE='" + Product_Code
				+ "') as A,"
				+ "(  SELECT coalesce(SUM(COALESCE(ORD_DTL.ORDERED_QTY,0)),0) AS ORDERED_QTY  FROM   MIGSRCMIL.ORDER as ORDR"
				+ " INNER JOIN MIGSRCMIL.ORDER_DTL as ORD_DTL ON ORDR.DLVRY_TICKET_SYSID = ORD_DTL.DLVRY_TICKET_SYSID "
				+ "LEFT JOIN MIGSRCMIL.V_LOT_HEADER as LH  ON LH.USCS_LOT_NUMBER = ORD_DTL.USCS_LOT_NUMBER "
				+ "LEFT JOIN MIGSRCMIL.DTL_CODE AS DT_MDE ON ORD_DTL.METHOD_OF_DTL_ENTRY_SYSID=DT_MDE.DTL_CODE_SYSID "
				+ "LEFT JOIN  MIGSRCMIL.DTL_CODE as DC2 ON ORD_DTL.BRAND_CODE_SYSID = DC2.DTL_CODE_SYSID "
				+ "LEFT JOIN MIGSRCMIL.DTL_CODE AS DT_ST ON ORDR.ORDER_STATUS_SYSID=DT_ST.DTL_CODE_SYSID "
				+ "INNER JOIN MIGSRCMIL.V_WHSE_CUST as WC ON ORDR.WHSE_SYSID = WC.WHSE_SYSID AND ORDR.CUST_SYSID = WC.CUST_SYSID "
				+ "INNER JOIN MIGSRCMIL.V_WHSE_PROD as VWP ON VWP.PROD_SYSID=ORD_DTL.PROD_SYSID AND VWP.WHSE_SYSID=ORD_DTL.WHSE_SYSID AND "
				+ "VWP.CUST_SYSID=ORD_DTL.CUST_SYSID "
				+ "WHERE  DT_MDE.DTL_CODE NOT IN ('PRD','CDT','PSQ')  AND DT_ST.DTL_CODE='0' AND " + " VWP.WHSE_SYSID="
				+ WHSE + " AND VWP.CUST_SYSID=" + CUST + " AND VWP.PROD_CODE='" + Product_Code + "') as B,"
				+ "(SELECT coalesce(SUM(COALESCE(ODL.CONVERTED_QTY,0)),0) AS C_ORDERED_QTY FROM   MIGSRCMIL.ORDER as ORDR "
				+ "INNER JOIN MIGSRCMIL.ORDER_DTL as ORD_DTL ON ORDR.DLVRY_TICKET_SYSID = ORD_DTL.DLVRY_TICKET_SYSID "
				+ "INNER JOIN MIGSRCMIL.ORDER_DTL_LOT as ODL ON ORD_DTL.ORDER_DTL_SYSID = ODL.ORDER_DTL_SYSID "
				+ "LEFT JOIN MIGSRCMIL.V_LOT_HEADER as LH  ON LH.USCS_LOT_NUMBER = ORD_DTL.USCS_LOT_NUMBER "
				+ "LEFT JOIN MIGSRCMIL.DTL_CODE AS DT_MDE ON ORD_DTL.METHOD_OF_DTL_ENTRY_SYSID=DT_MDE.DTL_CODE_SYSID "
				+ "LEFT JOIN MIGSRCMIL.DTL_CODE AS DT_ST ON ORDR.ORDER_STATUS_SYSID=DT_ST.DTL_CODE_SYSID "
				+ "INNER JOIN MIGSRCMIL.V_WHSE_CUST as WC ON ORDR.WHSE_SYSID = WC.WHSE_SYSID AND ORDR.CUST_SYSID = WC.CUST_SYSID "
				+ "INNER JOIN MIGSRCMIL.V_WHSE_PROD as VWP ON VWP.PROD_SYSID=ORD_DTL.PROD_SYSID AND VWP.WHSE_SYSID=ORD_DTL.WHSE_SYSID AND "
				+ "VWP.CUST_SYSID=ORD_DTL.CUST_SYSID "
				+ "WHERE  DT_MDE.DTL_CODE   IN ('PRD','CDT','PSQ') AND DT_ST.DTL_CODE='0' AND " + " VWP.WHSE_SYSID="
				+ WHSE + " AND VWP.CUST_SYSID=" + CUST + " AND VWP.PROD_CODE='" + Product_Code + "' ) as C";
		System.out.println(query); // 98403 98596
		ResultSet res = stmt.executeQuery(query);
		System.out.println("\n");

		if (res.next()) {
			//rowhead = sheet.createRow((short) 1);
			String Reserved_qty = res.getString("Reserved_Qty");
			rowhead.createCell(2).setCellValue(Reserved_qty);
			int rqty = Integer.parseInt(Reserved_qty.replace(",", ""));
			R_qty = rqty;
			System.out.println(String.format("%s -", res.getString(1)));
		}
		System.out.println("Reserved Qty:");
	}

	public void PIP_Qty(String WHSE, String CUST, String Product_Code) {
		try {
			
			String query = " "
					+ "SELECT CASE WHEN sum(PTD.QTY_TO_PICK) is null THEN '0' ELSE sum(PTD.QTY_TO_PICK) END as PIP_QTY FROM MIGSRCMIL.PICK_TASK_DTL as PTD "
					+ "INNER JOIN MIGSRCMIL.INVTRY_PALLET_DTL AS IPD ON PTD.INVTRY_PALLET_DTL_SYSID = IPD.INVTRY_PALLET_DTL_SYSID "
					+ "INNER JOIN MIGSRCMIL.V_LOT_HEADER AS VLH ON IPD.USCS_LOT_NUMBER_SYSID=VLH.USCS_LOT_NUMBER_SYSID "
					+ "INNER JOIN MIGSRCMIL.ORDER AS O ON PTD.DLVRY_TICKET_SYSID = O.DLVRY_TICKET_SYSID "
					+ "INNER JOIN MIGSRCMIL.ORDER_DTL as OD ON O.DLVRY_TICKET_SYSID = OD.DLVRY_TICKET_SYSID "
					+ "INNER JOIN MIGSRCMIL.DTL_CODE AS DC ON DC.DTL_CODE_SYSID=O.ORDER_STATUS_SYSID "
					+ "INNER JOIN MIGSRCMIL.V_WHSE_PROD as VWP ON VWP.PROD_SYSID=IPD.PROD_SYSID AND VWP.WHSE_SYSID=IPD.WHSE_SYSID AND "
					+ "VWP.CUST_SYSID=IPD.CUST_SYSID " + "WHERE VWP.WHSE_SYSID=" + WHSE + " AND VWP.CUST_SYSID=" + CUST
					+ " AND VWP.PROD_CODE='" + Product_Code + "' "
					+ "AND DC.DTL_CODE='1' AND PTD.PROJECTION_FLAG = 'N'";
			System.out.println(query); // 98403 98596
			ResultSet rs = stmt.executeQuery(query);
			System.out.println("\n");
			if (rs.next()) {
				String PIP_qty = rs.getString("PIP_QTY");
				rowhead.createCell(3).setCellValue(PIP_qty);
				P_qty = Integer.valueOf(PIP_qty);
				System.out.println(String.format("%s -", rs.getString(1)));
			}
			System.out.println("PIP Qty");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void Avail_Qty() throws SQLException, IOException {
		
		int A_qty = (T_qty - P_qty - H_qty - R_qty);
		String Avail_qty = String.format("%,d", A_qty);

		rowhead.createCell(4).setCellValue(Avail_qty);
		System.out.println(Avail_qty + " -");
		System.out.println("Avail Qty");
	}

	public void Hold_Qty(String WHSE, String CUST, String Product_Code) {
		try {
			
			String query = "SELECT CASE WHEN sum(IPD.ONHAND_QTY) is null THEN '0' ELSE TRIM(to_char(sum(IPD.ONHAND_QTY),'9,999')) END as Hold_Qty FROM MIGSRCMIL.PALLET_HOLD_CODE PHC "
					+ "INNER JOIN MIGSRCMIL.INVTRY_PALLET_DTL as IPD ON PHC.INVTRY_PALLET_DTL_SYSID = IPD.INVTRY_PALLET_DTL_SYSID "
					+ "INNER JOIN MIGSRCMIL.V_LOT_HEADER as VLH ON VLH.USCS_LOT_NUMBER_SYSID=IPD.USCS_LOT_NUMBER_SYSID "
					+ "INNER JOIN MIGSRCMIL.V_WHSE_PROD as VWP ON VWP.PROD_SYSID=IPD.PROD_SYSID AND VWP.WHSE_SYSID=IPD.WHSE_SYSID AND "
					+ "VWP.CUST_SYSID=IPD.CUST_SYSID " + "WHERE VWP.WHSE_SYSID=" + WHSE + " AND VWP.CUST_SYSID=" + CUST
					+ " AND VWP.PROD_CODE='" + Product_Code + "' ";
			System.out.println(query); // 98403 98596
			ResultSet res = stmt.executeQuery(query);
			System.out.println("\n");

			if (res.next()) {
				String Hold_qty = res.getString("Hold_Qty");
				rowhead.createCell(5).setCellValue(Hold_qty);
				int Hq = Integer.valueOf(Hold_qty.replace(",", ""));
				H_qty = Hq;
				System.out.println(String.format("%s -", res.getString(1)));
			}
			System.out.println("Hold Qty");

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void Total_Qty(String WHSE, String CUST, String Product_Code) {
		try {
			String query = "SELECT CASE WHEN sum(IPD.ONHAND_QTY) IS NULL THEN '0' ELSE TRIM(to_char(sum(IPD.ONHAND_QTY),'9,999')) END AS Total_QTY FROM MIGSRCMIL.INVTRY_PALLET_DTL AS IPD "
					+ "INNER JOIN MIGSRCMIL.V_LOT_HEADER AS VLH ON VLH.USCS_LOT_NUMBER_SYSID=IPD.USCS_LOT_NUMBER_SYSID "
					+ "INNER JOIN MIGSRCMIL.V_WHSE_PROD AS VWP ON VWP.PROD_SYSID=IPD.PROD_SYSID AND VWP.WHSE_SYSID=IPD.WHSE_SYSID AND "
					+ "VWP.CUST_SYSID=IPD.CUST_SYSID " + "WHERE VWP.WHSE_SYSID=" + WHSE + " AND VWP.CUST_SYSID=" + CUST
					+ " AND VWP.PROD_CODE='" + Product_Code + "' ";
			System.out.println(query); // 98403 98596
			ResultSet rs = stmt.executeQuery(query);
			System.out.println("\n");

			if (rs.next()) {
				String Total_qty = rs.getString("Total_QTY");
				rowhead.createCell(11).setCellValue(Total_qty);
				int Tq = Integer.valueOf(Total_qty.replace(",", ""));
				T_qty = Tq;
				System.out.println(String.format("%s -", rs.getString(1)));
			}
			System.out.println("Total Qty");

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void Total_Qty_Ord(String WHSE, String CUST, String Product_Code) {
		try {
			
			String query = "SELECT 	CASE WHEN SUM(OD.ORDERED_QTY) is NULL THEN '0' ELSE trim(TO_CHAR(SUM(OD.ORDERED_QTY),'9,999')) "
					+ "END as Total_ORDERED_QTY FROM MIGSRCMIL.ORDER as ORDR "
					+ "INNER JOIN MIGSRCMIL.ORDER_DTL as OD ON ORDR.DLVRY_TICKET_SYSID = OD.DLVRY_TICKET_SYSID "
					+ "INNER JOIN MIGSRCMIL.DTL_CODE as DC ON ORDR.ORDER_STATUS_SYSID = DC.DTL_CODE_SYSID AND DC.DTL_CODE NOT IN ('8','9','10') "
					+ "INNER JOIN MIGSRCMIL.V_WHSE_PROD as VWP ON VWP.PROD_SYSID=OD.PROD_SYSID AND VWP.WHSE_SYSID=OD.WHSE_SYSID AND VWP.CUST_SYSID=OD.CUST_SYSID "
					+ "WHERE VWP.WHSE_SYSID=" + WHSE + " AND VWP.CUST_SYSID=" + CUST + " AND VWP.PROD_CODE='"
					+ Product_Code + "' ";
			System.out.println(query); // 98403 98596
			ResultSet rs = stmt.executeQuery(query);
			System.out.println("\n");

			if (rs.next()) {
				String Total_qty_ord = rs.getString("Total_ORDERED_QTY");
				rowhead.createCell(6).setCellValue(Total_qty_ord);
				int Tq = Integer.valueOf(Total_qty_ord.replace(",", ""));
				O_qty = Tq;
				System.out.println(String.format("%s -", rs.getString(1)));
			}
			
			System.out.println("Total Qty Ordered");

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void Total_Qty_Avail_Ord() throws IOException {
		
		// DecimalFormat twoPlaces = new DecimalFormat("0,000");
		int T_qty_Ava_ord = (T_qty - P_qty - H_qty - R_qty - O_qty);
		String TT_Avail_qty = String.format("%,d", T_qty_Ava_ord);

		rowhead.createCell(7).setCellValue(TT_Avail_qty);
		System.out.println(TT_Avail_qty + " -");
		System.out.println("Total Qty Avail to Ordered");
	}

	public void Exp_Inb_Qty(String WHSE, String CUST, String Product_Code) {
		try {
			String query = "SELECT CASE when SUM(RC.XPCTD_QTY) is null THEN '0' ELSE TRIM(to_char(SUM(RC.XPCTD_QTY),'9,999')) END Exp_Inb_Qty "
					+ "FROM MIGSRCMIL.REC as R "
					+ "INNER JOIN MIGSRCMIL.REC_DTL AS RC ON R.REC_NUMBER_SYSID = RC.REC_NUMBER_SYSID "
					+ "LEFT JOIN MIGSRCMIL.DTL_CODE  AS DC ON RC.STATUS_SYSID  = DC.DTL_CODE_SYSID "
					+ "INNER JOIN MIGSRCMIL.V_WHSE_PROD  AS VWP ON VWP.PROD_SYSID=RC.PROD_SYSID AND VWP.WHSE_SYSID=RC.WHSE_SYSID AND "
					+ "VWP.CUST_SYSID=RC.CUST_SYSID " + "WHERE DC.DTL_CODE ='P' AND VWP.WHSE_SYSID=" + WHSE
					+ " AND VWP.CUST_SYSID=" + CUST + " AND VWP.PROD_CODE='" + Product_Code + "' ";
			System.out.println(query); // 98403 98596
			ResultSet rs = stmt.executeQuery(query);
			System.out.println("\n");

			if (rs.next()) {
				String Total_qty_ord = rs.getString("Exp_Inb_Qty");
				rowhead.createCell(10).setCellValue(Total_qty_ord);
				System.out.println(String.format("%s -", rs.getString(1)));
			}
			System.out.println("Expeceted Inbound Qty");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void Potential_Short_Qty() throws IOException {
		
		int Pot_qty = (T_qty - P_qty - H_qty - R_qty);
		String Potential_qty = String.format("%,d", Pot_qty);
		if (Pot_qty < 0) {
			rowhead.createCell(8).setCellValue(Potential_qty);
			System.out.println(Potential_qty + " -");
		} else {
			rowhead.createCell(8).setCellValue("0");
			System.out.println("0" + " -");
		}
		System.out.println("Potential Short Qty");

	}

	public void Rec_In_Proc_Qty(String WHSE, String CUST, String Product_Code) {
		try {
			
			String query = " SELECT CASE WHEN SUM(RC.XPCTD_QTY) is null THEN '0' ELSE trim(to_char(sum(RC.XPCTD_QTY),'9,999')) "
					+ "END AS Rec_In_Proc_Qty FROM MIGSRCMIL.REC AS R "
					+ "INNER JOIN MIGSRCMIL.REC_DTL AS RC ON RC.REC_NUMBER_SYSID = R.REC_NUMBER_SYSID "
					+ "LEFT JOIN MIGSRCMIL.DTL_CODE  AS DC ON R.RECEIPT_STATUS_SYSID  = DC.DTL_CODE_SYSID "
					+ "INNER JOIN MIGSRCMIL.V_WHSE_PROD AS VWP ON VWP.PROD_SYSID=RC.PROD_SYSID AND VWP.WHSE_SYSID=RC.WHSE_SYSID AND "
					+ "VWP.CUST_SYSID=RC.CUST_SYSID " + "WHERE  DC.DTL_CODE ='W' AND VWP.WHSE_SYSID=" + WHSE
					+ " AND VWP.CUST_SYSID=" + CUST + " AND VWP.PROD_CODE='" + Product_Code + "' ";
			System.out.println(query); // 98403 98596
			ResultSet rs = stmt.executeQuery(query);
			System.out.println("\n");

			if (rs.next()) {
				String Rec_In_Proc_Qty = rs.getString("Rec_In_Proc_Qty");
				rowhead.createCell(9).setCellValue(Rec_In_Proc_Qty);
				System.out.println(String.format("%s -", rs.getString(1)));
			}
			
			System.out.println("Receipt in process Qty");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	@Test(dependsOnMethods="excelCreation",dataProvider = "exceldata")
	public void Inv_Query(String rowno,String Whse_Sysid,String Cust_Sysid,String Prd) throws Exception {
		int rowCount = Integer.parseInt(rowno);
		PRD(rowCount,Whse_Sysid, Cust_Sysid, Prd);
		Reserved_Qty(Whse_Sysid, Cust_Sysid, Prd);
		PIP_Qty(Whse_Sysid, Cust_Sysid, Prd);
		Hold_Qty(Whse_Sysid, Cust_Sysid, Prd);
		Total_Qty(Whse_Sysid, Cust_Sysid, Prd);
		Avail_Qty();
		Total_Qty_Ord(Whse_Sysid, Cust_Sysid, Prd);
		Total_Qty_Avail_Ord();
		Potential_Short_Qty();
		Rec_In_Proc_Qty(Whse_Sysid, Cust_Sysid, Prd);
		Exp_Inb_Qty(Whse_Sysid, Cust_Sysid, Prd);
	}
	
	@Test(dependsOnMethods = "Inv_Query")
	public void eUSCOLD_Login() throws Exception {
		WB.write(fileout);
		fileout.close();
		System.out.println("Your excel file has been generated!");
		login();
	}

	@Test(dependsOnMethods = "eUSCOLD_Login")
	public void Inventory() throws Exception {
		driver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);
		Actions a = new Actions(driver);
		WebElement w = driver.findElement(By.cssSelector("#inv"));
		a.moveToElement(w).perform();
		Thread.sleep(1000);
		WebElement w1 = driver.findElement(By.cssSelector("#inv1>div"));
		a.moveToElement(w1).perform();
		Thread.sleep(1000);
		WebElement w2 = driver.findElement(By.cssSelector("#inv11>div"));
		a.moveToElement(w2).click().perform();
		Thread.sleep(6000);
	}

	@Test(dependsOnMethods = "Inventory", dataProvider = "exceldata")
	public void prod(String Product, String ProductDesc,String ReservedQty,String PickPrcQty,String AvailQty,String HoldQty,
			String TotalQtyOrd,String TotalQtyAvailOrd,String Pot_Short_Qty,String RIPQty,String ExpINBQty,String TotalQty) {
		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		try {
			new Select(driver.findElement(By.xpath("//*[@id='customerNames']"))).deselectByValue("ALL");
			Thread.sleep(1000);
			new Select(driver.findElement(By.xpath("//*[@id='customerNames']"))).selectByValue(Cust_Number); // 0616016030
			Thread.sleep(1000);
			driver.findElement(By.xpath("//*[@id='productCode']")).clear(); // 511600000100950
			Thread.sleep(1000);
			if(!Product.equals("")) {
			driver.findElement(By.xpath("//*[@id='productCode']")).sendKeys(Product);
			Thread.sleep(1000);
			driver.findElement(By.xpath("//td[1][@valign='top']/a/img")).click();
			Thread.sleep(1000);
			driver.findElement(By.xpath("//a[@class='btnGray']/span")).click();
			Thread.sleep(8000);
			Data = driver.findElement(By.xpath("//td[@class='bdyBkgInner']")).getText();
			if(Data.contains("There is no data available") && !Product.equals("")){
				String data1 = driver.findElement(By.xpath("//*[@id='row']/tbody/tr/td")).getText();
				log.error("There is no data available in UI$$" + data1 + "$$ $$" + "Data is Available in DB" + "$$");
				driver.findElement(By.xpath("//td[@class='brdcrmText']/a")).click();
				Thread.sleep(5000);
			}else if(!Data.contains("There is no data available") && Product.equals(null)){
				String data1 = driver.findElement(By.xpath("//*[@id='row']/tbody/tr/td")).getText();
				log.error("Data available in UI$$" + data1 + "$$ $$" + "There is no data available in DB" + "$$");
				driver.findElement(By.xpath("//td[@class='brdcrmText']/a")).click();
				Thread.sleep(5000);
			}else  {
				String Prd = driver.findElement(By.xpath("//*[@id='row']/tbody/tr/td[1]")).getText();
				if (Prd.equals(Product)) {
					log.info("Product Code$$" + Prd + "$$Product$$" + Product + "$$");
				} else {
					log.error("Product Code$$" + Prd + "$$Product$$" + Product + "$$");
				}
				String PrdDesc = driver.findElement(By.xpath("//*[@id='row']/tbody/tr/td[2]")).getText();
				if (PrdDesc.equals(ProductDesc)) {
					log.info("Product Description$$" + PrdDesc + "$$Product Description$$" + ProductDesc + "$$");
				} else {
					log.error("Product Description$$" + PrdDesc + "$$Product Description$$" + ProductDesc + "$$");
				}
				String ResQty = driver.findElement(By.xpath("//*[@id='row']/tbody/tr/td[3]")).getText();
				if (ResQty.equals(ReservedQty)) {
					log.info("Reserved Qty$$" + ResQty + "$$Reserved Qty$$" + ReservedQty + "$$");
				} else {
					log.error("Reserved Qty$$" + ResQty + "$$Reserved Qty$$" + ReservedQty + "$$");
				}
				String PIPQty = driver.findElement(By.xpath("//*[@id='row']/tbody/tr/td[4]")).getText();
				if (PIPQty.equals(PickPrcQty)) {
					log.info("PIP Qty$$" + PIPQty + "$$PIP Qty$$" + PickPrcQty + "$$");
				} else {
					log.error("PIP Qty$$" + PIPQty + "$$PIP Qty$$" + PickPrcQty + "$$");
				}
				String AvaQty = driver.findElement(By.xpath("//*[@id='row']/tbody/tr/td[5]")).getText();
				if (AvaQty.equals(AvailQty)) {
					log.info("Avail Qty$$" + AvaQty + "$$Avail Qty$$" + AvailQty + "$$");
				} else {
					log.error("Avail Qty$$" + AvaQty + "$$Avail Qty$$" + AvailQty + "$$");
				}
				String HldQty = driver.findElement(By.xpath("//*[@id='row']/tbody/tr/td[6]")).getText();
				if (HldQty.equals(HoldQty)) {
					log.info("Hold Qty$$" + HldQty + "$$Hold Qty$$" + HoldQty + "$$");
				} else {
					log.error("Hold Qty$$" + HldQty + "$$Hold Qty$$" + HoldQty + "$$");
				}
				String TQtyOrd = driver.findElement(By.xpath("//*[@id='row']/tbody/tr/td[7]")).getText();
				if (TQtyOrd.equals(TotalQtyOrd)) {
					log.info("Total Qty Ordered$$" + TQtyOrd + "$$Total Qty Ordered$$" + TotalQtyOrd + "$$");
				} else {
					log.error("Total Qty Ordered$$" + TQtyOrd + "$$Total Qty Ordered$$" + TotalQtyOrd + "$$");
				}
				String TQtyAvOrd = driver.findElement(By.xpath("//*[@id='row']/tbody/tr/td[8]")).getText();
				if (TQtyAvOrd.equals(TotalQtyAvailOrd)) {
					log.info("Total Qty Avail to Order$$" + TQtyAvOrd + "$$Total Qty Avail to Order$$"
							+ TotalQtyAvailOrd + "$$");
				} else {
					log.error("Total Qty Avail to Order$$" + TQtyAvOrd + "$$Total Qty Avail to Order$$"
							+ TotalQtyAvailOrd + "$$");
				}
				String PotQty = driver.findElement(By.xpath("//*[@id='row']/tbody/tr/td[9]")).getText();
				if (PotQty.equals(Pot_Short_Qty)) {
					log.info("Potential Short Qty$$" + PotQty + "$$Potential Short Qty$$" + Pot_Short_Qty + "$$");
				} else {
					log.error("Potential Short Qty$$" + PotQty + "$$Potential Short Qty$$" + Pot_Short_Qty + "$$");
				}
				String RecinPrcQty = driver.findElement(By.xpath("//*[@id='row']/tbody/tr/td[10]")).getText();
				if (RecinPrcQty.equals(RIPQty)) {
					log.info("Recipt In Process Qty$$" + RecinPrcQty + "$$Recipt In Process Qty$$" + RIPQty + "$$");
				} else {
					log.error("Recipt In Process Qty$$" + RecinPrcQty + "$$Recipt In Process Qty$$" + RIPQty + "$$");
				}
				String ExpInbQty = driver.findElement(By.xpath("//*[@id='row']/tbody/tr/td[11]")).getText();
				if (ExpInbQty.equals(ExpINBQty)) {
					log.info("Expected Inbound Qty$$" + ExpInbQty + "$$Expected Inbound Qty$$" + ExpINBQty + "$$");
				} else {
					log.error("Expected Inbound Qty$$" + ExpInbQty + "$$Expected Inbound Qty$$" + ExpINBQty + "$$");
				}
				String TtlQty = driver.findElement(By.xpath("//*[@id='row']/tbody/tr/td[12]")).getText();
				if (TtlQty.equals(TotalQty)) {
					log.info("Total Qty$$" + TtlQty + "$$Total Qty$$" + TotalQty + "$$");
				} else {
					log.error("Total Qty$$" + TtlQty + "$$Total Qty$$" + TotalQty + "$$");
				}
				Thread.sleep(2000);
				driver.findElement(By.xpath("//td[@class='brdcrmText']/a")).click();
				Thread.sleep(5000);				
			}
			}else {
				log.info("Data$$" + "There is no data available in UI" + "$$ $$" + "There is no data available in DB" + "$$");
				driver.navigate().refresh();
				Thread.sleep(5000);
			}
		} catch (Exception e) {
			System.out.println("Exception is :" + e.getMessage());
		}
	}
	
	@AfterTest
	public void logout() throws Exception{
		driver.findElement(By.linkText("Logout")).click();
		Thread.sleep(2000);
		driver.close();
	}

	@DataProvider(name = "exceldata")
	public Object[][] Readexcel(Method m) throws IOException, BiffException {
		FileInputStream fi = null;
		Workbook w;
		Sheet s =null;
		if(m.getName().equals("Inv_Query")) {
			fi = new FileInputStream(Path+"DB_Input_Details.xls");
			w = Workbook.getWorkbook(fi);
			s = w.getSheet("Inv_Inq_Prod");
		}else if(m.getName().equals("prod")){
			fi = new FileInputStream(test_results+"Inventory_Inquiry_Prod.xls");
			w = Workbook.getWorkbook(fi);
			s = w.getSheet("Product");
		}		 
		int Rows = s.getRows() - 1;
		int Columns = s.getColumns();
		String InputData[][] = new String[Rows][Columns];
		for (int i = 1; i <= Rows; i++) {
			for (int j = 0; j < Columns; j++) {
				Cell c = s.getCell(j, i);
				InputData[i - 1][j] = c.getContents();
				//System.out.println(InputData[i - 1][j]);
			}
		}
		return InputData;
	}
}