package com.euscold.testcases;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Method;
import java.sql.Connection;
import java.sql.ResultSet;
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

public class Order_Summary_Multiple_OrderNos extends DBConnection {


	
	public void Order_No(int rowno, String WHSE, String CUST, String Ord_No,String from_date,String to_date) {
		try {
			rowhead = sheet.createRow(0);
			rowhead.createCell(0).setCellValue("Order Number");
			rowhead.createCell(1).setCellValue("Expected ship date");
			rowhead.createCell(2).setCellValue("Delivery Ticket");
			rowhead.createCell(3).setCellValue("PONo");
			rowhead.createCell(4).setCellValue("Load Number");
			rowhead.createCell(5).setCellValue("Mob No");
			rowhead.createCell(6).setCellValue("Consignee Name");
			rowhead.createCell(7).setCellValue("Qty");
			rowhead.createCell(8).setCellValue("Status");
			datasource.setUser("HQIRCHEN");
			datasource.setPassword("HQIRCHEN");
			Connection connection = datasource.getConnection();
			stmt = connection.createStatement();
			String query = " SELECT TRIM(ORDER_NUMBER) ORD,VARCHAR_FORMAT(XPCTD_SHIP_DATE,'mm/dd/yyyy') Shipt_date,TRIM(DLVRY_TICKET_NUMBER) DT,"
					+ "TRIM(PO_NUMBER) PO,LOAD_NUMBER,TRIM(MASTER_LINK_NUMBER) MOB,CONSIGNEE_NAME,"
					+ "TRIM(to_char(sum(ORDERED_QTY), '9,999')) QTY,CASE WHEN MIGSRCMIL.DTL_CODE.CODE_DESC = 'Posted' THEN 'Shipped' "
					+ "WHEN MIGSRCMIL.DTL_CODE.CODE_DESC = 'Open' THEN 'Pending' END AS CODE_DESC FROM MIGSRCMIL.ORDER_DTL "
					+ "INNER JOIN  MIGSRCMIL.ORDER ON MIGSRCMIL.ORDER.DLVRY_TICKET_SYSID = MIGSRCMIL.ORDER_DTL.DLVRY_TICKET_SYSID "
					+ "INNER JOIN MIGSRCMIL.V_CONSIGNEE ON  MIGSRCMIL.V_CONSIGNEE.CONS_SYSID=MIGSRCMIL.ORDER.CONS_SYSID "
					+ "INNER JOIN MIGSRCMIL.DTL_CODE    ON MIGSRCMIL.DTL_CODE.DTL_CODE_SYSID=MIGSRCMIL.ORDER.ORDER_STATUS_SYSID "
					+ " WHERE MIGSRCMIL.ORDER.WHSE_SYSID=" + WHSE + " AND  MIGSRCMIL.ORDER.CUST_SYSID=" + CUST
					+ " AND MIGSRCMIL.ORDER.ORDER_NUMBER='" + Ord_No + "' AND varchar_format(MIGSRCMIL.ORDER.XPCTD_SHIP_DATE,'YYYYMMDD') BETWEEN '"+from_date+"' AND '"+to_date+"'"
					+ "GROUP BY ORDER_NUMBER,XPCTD_SHIP_DATE,DLVRY_TICKET_NUMBER,PO_NUMBER,LOAD_NUMBER,MASTER_LINK_NUMBER,CONSIGNEE_NAME,CODE_DESC";
			System.out.println(query);
			ResultSet res = stmt.executeQuery(query);
			System.out.println("\n");
			if (res.next()) {
				rowhead = sheet.createRow((short) rowno);
				rowhead.createCell(0).setCellValue(res.getString(("ORD")));
				rowhead.createCell(1).setCellValue(res.getString("Shipt_date"));
				rowhead.createCell(2).setCellValue(res.getString("DT"));
				rowhead.createCell(3).setCellValue(res.getString("PO"));
				rowhead.createCell(4).setCellValue(res.getString("LOAD_NUMBER"));
				rowhead.createCell(5).setCellValue(res.getString("MOB"));
				rowhead.createCell(6).setCellValue(res.getString("CONSIGNEE_NAME"));
				rowhead.createCell(7).setCellValue(res.getString("QTY"));
				rowhead.createCell(8).setCellValue(res.getString("CODE_DESC"));
				System.out.println(String.format("%s - %s - %s - %s - %s - %s - %s - %s", res.getString(1),
						res.getString(2), res.getString(3), res.getString(4), res.getString(5), res.getString(6),
						res.getString(7), res.getString(8)));
			} else {
				rowhead = sheet.createRow((short) rowno);
				No_data = "No Data Available in DB";
			}		
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
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
	public void excelCreation() throws FileNotFoundException {
		fileout = new FileOutputStream(test_results+"Order_Summary.xls");
		sheet = WB.createSheet("OrderNo");		
	}
	
	@Test(dependsOnMethods="excelCreation",dataProvider="exceldata")
	public void order_inq(String rowno,String whse_sysid,String cust_sysid,String ord_no,String from_date
			,String to_date) {
		int rowCount = Integer.parseInt(rowno);
		Order_No(rowCount,whse_sysid, cust_sysid, ord_no,from_date,to_date);
	}

	@Test(dependsOnMethods = "order_inq")
	public void eUSCOLD_Login() throws Exception {
		WB.write(fileout);
		fileout.close();
		System.out.println("Your excel file has been generated!");
		login();
	}

	@Test(dependsOnMethods = "eUSCOLD_Login")
	public void order_search() throws Exception {
		driver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);
		Actions a = new Actions(driver);
		WebElement w = driver.findElement(By.cssSelector("#order"));
		a.moveToElement(w).perform();
		Thread.sleep(1000);
		WebElement w1 = driver.findElement(By.cssSelector("#order2"));
		a.moveToElement(w1).perform();
		Thread.sleep(1000);
		WebElement w2 = driver.findElement(By.cssSelector("#order21>div"));
		a.moveToElement(w2).click().perform();
		Thread.sleep(5000);
	}

	@Test(dependsOnMethods = "order_search", dataProvider = "exceldata")
	public void Ord_Summary(String OrderNumber, String shpDt, String TicketNo, String PONo, String LoadNo, String MobNo,
			String Consignee, String Order_Qty, String Status) throws Exception {
		try {
			driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
			new Select(driver.findElement(By.xpath("//*[@id='customerNames']"))).deselectByVisibleText("ALL");
			Thread.sleep(1000);
			new Select(driver.findElement(By.xpath("//*[@id='customerNames']"))).selectByValue(Cust_Number);
			Thread.sleep(1000);
			driver.findElement(By.cssSelector("#tempfromDate")).clear();
			driver.findElement(By.cssSelector("#tempfromDate")).sendKeys(UI_from_date);
			driver.findElement(By.cssSelector("#temptoDate")).clear();
			driver.findElement(By.cssSelector("#temptoDate")).sendKeys(UI_to_date);
			Thread.sleep(1000);
			driver.findElement(By.xpath("//*[@id='custOrderNumber']")).click();
			driver.findElement(By.xpath("//*[@id='custOrderNumber']")).sendKeys(OrderNumber);
			Thread.sleep(1000);
			driver.findElement(By.xpath("//a[@class='btnGray']/span/label")).click();
			Thread.sleep(5000);
			String Data = driver.findElement(By.xpath("//div[@id='tabCont1']")).getText();
			if (Data.contains("There is no data available") && OrderNumber.equals(null)) {
				String data1 = driver.findElement(By.xpath("//*[@id='row']/tbody/tr/td")).getText().trim();
				log.info("No Data Available in UI$$" + data1 + "$$ $$" + No_data + "$$");
				Thread.sleep(2000);
			} else if (Data.contains("There is no data available") && !OrderNumber.equals(null)) {
				String data1 = driver.findElement(By.xpath("//*[@id='row']/tbody/tr/td")).getText().trim();
				log.info("No Data Available in UI$$" + data1 + "$$ $$" + "Data is Available in DB" + "$$");
				Thread.sleep(2000);
			} else {
				for (int i = 1; i <= 9; i++) {
					String No = Integer.toString(i);
					if (No.equals("1")) {
						String OrderNo = driver.findElement(By.xpath("//*[@id='row']/tbody/tr/td[" + i + "]"))
								.getText().trim();
						if (OrderNo.equals(OrderNumber)) {
							log.info("OrderNumber$$" + OrderNo + "$$OrderNumber$$" + OrderNumber + "$$");
						} else {
							log.error("OrderNumber$$" + OrderNo + "$$OrderNumber$$" + OrderNumber + "$$");
						}
					} else if (No.equals("2")) {
						String ShipDate = driver.findElement(By.xpath("//*[@id='row']/tbody/tr/td[" + i + "]"))
								.getText().trim();
						if (ShipDate.equals(shpDt)) {
							log.info("ShipDate$$" + ShipDate + "$$ShipDate$$" + shpDt + "$$");
						} else {
							log.error("ShipDate$$" + ShipDate + "$$ShipDate$$" + shpDt + "$$");
						}
					} else if (No.equals("3")) {
						String Ticket = driver.findElement(By.xpath("//*[@id='row']/tbody/tr/td[" + i + "]"))
								.getText().trim();
						String Tkt[] = Ticket.split(" ");
						if (Ticket.equals(TicketNo)) {
							log.info("TicketNo$$" + Tkt[0] + "$$TicketNo$$" + TicketNo + "$$");
						} else {
							log.error("TicketNo$$" + Tkt[0] + "$$TicketNo$$" + TicketNo + "$$");
						}
					} else if (No.equals("4")) {
						String PoNo = driver.findElement(By.xpath("//*[@id='row']/tbody/tr/td[" + i + "]")).getText().trim();
						if (PoNo.equals(PONo)) {
							log.info("PONo$$" + PoNo + "$$PONo$$" + PONo + "$$");
						} else {
							log.error("PONo$$" + PoNo + "$$PONo$$" + PONo + "$$");
						}
					} else if (No.equals("5")) {
						String LdNo = driver.findElement(By.xpath("//*[@id='row']/tbody/tr/td[" + i + "]")).getText().trim();
						if (LdNo.equals(LoadNo)) {
							log.info("LoadNo$$" + LdNo + "$$LoadNo$$" + LoadNo + "$$");
						} else {
							log.error("LoadNo$$" + LdNo + "$$LoadNo$$" + LoadNo + "$$");
						}
					} else if (No.equals("6")) {
						String MblNo = driver.findElement(By.xpath("//*[@id='row']/tbody/tr/td[" + i + "]")).getText().trim();
						if (MblNo.equals(MobNo)) {
							log.info("MobileNo$$" + MblNo + "$$MobileNo$$" + MobNo + "$$");
						} else {
							log.error("MobileNo$$" + MblNo + "$$MobileNo$$" + MobNo + "$$");
						}
					} else if (No.equals("7")) {
						String Consigne = driver.findElement(By.xpath("//*[@id='row']/tbody/tr/td[" + i + "]"))
								.getText().trim();
						if (Consigne.equals(Consignee)) {
							log.info("Consignee$$" + Consigne + "$$Consignee$$" + Consignee + "$$");
						} else {
							log.error("Consignee$$" + Consigne + "$$Consignee$$" + Consignee + "$$");
						}
					} else if (No.equals("8")) {
						String Qty = driver.findElement(By.xpath("//*[@id='row']/tbody/tr/td[" + i + "]")).getText().trim();
						if (Qty.equals(Order_Qty)) {
							log.info("Ordered Qty$$" + Qty + "$$Ordered Qty$$" + Order_Qty + "$$");
						} else {
							log.error("Ordered Qty$$" + Qty + "$$Ordered Qty$$" + Order_Qty + "$$");
						}
					} else {
						String status = driver.findElement(By.xpath("//*[@id='row']/tbody/tr/td[" + i + "]")).getText().trim();
						if (status.equals(Status)) {
							log.info("Status$$" + status + "$$Status$$" + Status + "$$");
						} else {
							log.error("Status$$" + status + "$$Status$$" + Status + "$$");
						}
					}
					Thread.sleep(1000);
				}
			}			
			driver.findElement(By.xpath("//td[@class='brdcrmText']/a")).click();
			Thread.sleep(3000);
		} catch (Exception e) {
			System.out.println("Exception is:" + e);
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
		if(m.getName().equals("order_inq")){
			fi = new FileInputStream(Path+"DB_Input_Details.xls");
			w = Workbook.getWorkbook(fi);
			s = w.getSheet("Ord_Inq");
		}else if(m.getName().equals("Ord_Summary")){
			fi = new FileInputStream(test_results+"Order_Summary.xls");
			w = Workbook.getWorkbook(fi);
			s = w.getSheet("OrderNo");
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
