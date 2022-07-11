import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;

import javax.swing.JOptionPane;

import org.apache.poi.hslf.usermodel.HSLFPictureData;
import org.apache.poi.hslf.usermodel.HSLFPictureShape;
import org.apache.poi.hslf.usermodel.HSLFSlide;
import org.apache.poi.hslf.usermodel.HSLFSlideShow;
import org.apache.poi.hslf.usermodel.HSLFTextBox;
import org.apache.poi.hslf.usermodel.HSLFTextParagraph;
import org.apache.poi.hslf.usermodel.HSLFTextRun;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFPictureData;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.sl.usermodel.TextParagraph.TextAlign;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.apache.poi.ss.usermodel.Font;

import java.awt.Color;
import java.awt.Dimension;

import java.awt.Rectangle;
import java.net.HttpURLConnection;
import java.net.URL;

import com.asana.Client;
import com.asana.dispatcher.AccessTokenDispatcher;
import com.asana.dispatcher.BasicAuthDispatcher;
import com.asana.dispatcher.Dispatcher;
import com.asana.dispatcher.OAuthDispatcher;
import com.asana.errors.AsanaError;
import com.asana.errors.RateLimitEnforcedError;
import com.asana.errors.RetryableAsanaError;
import com.asana.models.ResultBody;
import com.asana.models.Task;
import com.asana.models.Team;
import com.asana.models.User;
import com.asana.models.Workspace;
import com.asana.requests.Request;
import com.asana.resources.*;
import com.asana.resources.gen.TasksBase;
import com.google.api.client.http.*;
import com.google.api.client.json.webtoken.JsonWebSignature.Parser;
import com.google.common.base.Joiner;
import com.google.gson.JsonElement;
import com.google.gson.JsonObject;
import com.google.gson.JsonParser;

import java.awt.Rectangle;
import java.awt.geom.Rectangle2D;

import org.apache.poi.sl.usermodel.PictureData;
import org.apache.poi.sl.usermodel.TextParagraph.TextAlign;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
//import org.apache.poi.xslf.usermodel.XSLFPropertiesDelegate.XSLFFillProperties;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.css.RGBColor;
import org.apache.poi.hslf.usermodel.HSLFSlide;
import org.apache.poi.hslf.usermodel.HSLFSlideShow;
import org.apache.poi.hslf.usermodel.HSLFTextBox;
import org.apache.poi.hslf.usermodel.HSLFTextRun;

import java.awt.Color;
import java.awt.Dimension;
import java.awt.Rectangle;

import org.apache.commons.math3.util.MultidimensionalCounter.Iterator;
import org.apache.poi.*;

import org.apache.poi.hslf.model.*;
import org.apache.poi.hslf.usermodel.HSLFFill;
import org.apache.poi.hslf.usermodel.HSLFPictureData;
import org.apache.poi.hslf.usermodel.HSLFPictureShape;
import org.apache.poi.hslf.usermodel.HSLFSlide;
import org.apache.poi.hslf.usermodel.HSLFSlideShow;
import org.apache.poi.hslf.usermodel.HSLFTextBox;
import org.apache.poi.hslf.usermodel.HSLFTextParagraph;
import org.apache.poi.hslf.usermodel.HSLFTextRun;
import org.apache.poi.hslf.usermodel.HSLFTextShape;

import java.io.*;
import java.awt.Color;





public class Main {
	
	public enum rep{
		//Workspace gid:510943893270416
		//users api endpoint: 
		//https://app.asana.com/api/1.0/workspaces/510943893270416/users?opt_pretty
		
		Aidan,//1155943241735889
		Fulu,//1143898803372993
		Hazel,//662072092829294
		Jackie,//662072092829271
		Jolene,//1112174143264743
		Karen,//1163017665830673
		Kaylen,//1119664339723870
		William,//1138380100468293
		Louis,//804048676446699
		Mohamed,//1128742066746577
		Njabulo,//1113516810972893
		Roberto,//1164595080703813
		Roebeth,//1152557295298594
		Roland,//618873806134395
		Sipho,//1133859561713096
		Thabo,//1108363137349882
		KiYano,//1196557710846061
		Lindiwe,//1200006066577647
		Natasha,//1200728660300272
		Jimmy,//1201166032790665
		Aaron,//1201684458943755
		Campbell,//1201670713789795
		Lynette,//1201785607352422
	}

	private void powerPoint()
	{
		 //locates images
		 
		 File image = new File("C:\\Users\\Nick\\Desktop\\BlueMango\\POWERPOINT\\bin\\blueMangoLogo.png");
		 File backg = new File("C:\\Users\\Nick\\Desktop\\BlueMango\\POWERPOINT\\bin\\blueMangoBG.jpg");
		
		 File LHS = new File("C:\\Users\\Nick\\Desktop\\BlueMango\\POWERPOINT\\bin\\bluePanelLHS.png");
		 File copyright = new File("C:\\Users\\Nick\\Desktop\\BlueMango\\POWERPOINT\\bin\\blueMangoCopyRight.png");
		 File yellowborder = new File("C:\\Users\\Nick\\Desktop\\BlueMango\\POWERPOINT\\bin\\yellowBorder.png");
		
		 File VENDORS = new File("G:\\My Drive\\Sales Shared\\Principle Photo Upload\\Fulu-Free State\\Photos");
		 
		 
	}
	
	public static void main(String[] args) throws IOException, InvalidFormatException {
		// TODO Auto-generated method stub
		
		
		 String month = JOptionPane.showInputDialog("Reports for Month of? (numeric)"); 
		 File BMRep = new File("G:\\My Drive\\Sales Shared\\4.Principle Photo Upload & KPI\\KPI");
		
		 
		 
		 //TODO: add font stlye to headers
		 
		 
		 for (int i = 0; i < BMRep.list().length; i++) {
			
//			 if(BMRep.list()[i].compareToIgnoreCase("Lindiwe-Vaal, Botswana") != 0)
//				 continue;
			 
			 if(BMRep.list()[i].compareToIgnoreCase("desktop.ini") == 0)
				 continue;
			 
			 System.out.println(BMRep.list()[i]);
			 
			 //START OF KPI
				
			 //KRONOS LOCATION
					Workbook kronos = WorkbookFactory.create(new FileInputStream("G:\\My Drive\\Sales Shared\\9.Kronos Reports\\22Jan.xls"));
			 		//BM TARGETS LOCATION
				InputStream tfile = new FileInputStream("G:\\My Drive\\KPI\\KPI Monthly Data\\2022\\Apr22\\BM Targets.xls");
				
					//DORMANT FILE LOCATION
				Workbook dormant = WorkbookFactory.create(new FileInputStream("G:\\My Drive\\KPI\\KPI Monthly Data\\2022\\Apr22\\Reps Dormant Stock 20220519 for Apr22.xls"));
				
			 Workbook master = new HSSFWorkbook();
			 
			 CreationHelper createHelper = master.getCreationHelper();
			 
			 
			 CellStyle header = master.createCellStyle();
			 header.setWrapText(true);
			 Font font = master.createFont();
			 font.setBold(true);
			 
			 Font fred = master.createFont();
			 fred.setColor(fred.COLOR_RED);
			 
			 header.setBorderTop(BorderStyle.THIN);
			 header.setTopBorderColor(IndexedColors.BLACK.getIndex());
				
			 header.setBorderLeft(BorderStyle.THIN);
			 header.setLeftBorderColor(IndexedColors.BLACK.getIndex());
				
			 header.setBorderBottom(BorderStyle.THIN);
			 header.setBottomBorderColor(IndexedColors.BLACK.getIndex());
				
			 header.setBorderRight(BorderStyle.THIN);
			 header.setRightBorderColor(IndexedColors.BLACK.getIndex());
			 header.setFont(font);
			
			 Font tempfont = master.createFont();
				tempfont.setBold(false);
				
				//currency
				CellStyle style1 = master.createCellStyle();
				style1.setDataFormat(createHelper.createDataFormat().getFormat("R#,##0.00"));
				style1.setBorderTop(BorderStyle.THIN);
				style1.setTopBorderColor(IndexedColors.BLACK.getIndex());
				
				style1.setBorderLeft(BorderStyle.THIN);
				style1.setLeftBorderColor(IndexedColors.BLACK.getIndex());
				
				style1.setBorderBottom(BorderStyle.THIN);
				style1.setBottomBorderColor(IndexedColors.BLACK.getIndex());
				
				style1.setBorderRight(BorderStyle.THIN);
				style1.setRightBorderColor(IndexedColors.BLACK.getIndex());
				
				CellStyle style1f = master.createCellStyle();
				style1f.setDataFormat(createHelper.createDataFormat().getFormat("R#,##0.00"));
				style1f.setBorderTop(BorderStyle.THIN);
				style1f.setTopBorderColor(IndexedColors.BLACK.getIndex());
				
				style1f.setBorderLeft(BorderStyle.THIN);
				style1f.setLeftBorderColor(IndexedColors.BLACK.getIndex());
				
				style1f.setBorderBottom(BorderStyle.THIN);
				style1f.setBottomBorderColor(IndexedColors.BLACK.getIndex());
				
				style1f.setBorderRight(BorderStyle.THIN);
				style1f.setRightBorderColor(IndexedColors.BLACK.getIndex());
				
				style1f.setFont(fred);
				
				CellStyle style1b = master.createCellStyle();
				style1b.setDataFormat(createHelper.createDataFormat().getFormat("R#,##0.00"));
				style1b.setBorderTop(BorderStyle.THIN);
				style1b.setTopBorderColor(IndexedColors.BLACK.getIndex());
				
				style1b.setBorderLeft(BorderStyle.THIN);
				style1b.setLeftBorderColor(IndexedColors.BLACK.getIndex());
				
				style1b.setBorderBottom(BorderStyle.THIN);
				style1b.setBottomBorderColor(IndexedColors.BLACK.getIndex());
				
				style1b.setBorderRight(BorderStyle.THIN);
				style1b.setRightBorderColor(IndexedColors.BLACK.getIndex());
				style1b.setFont(font);
				
				//percentage
				CellStyle style2 = master.createCellStyle();
				style2.setDataFormat(createHelper.createDataFormat().getFormat("0.0%"));
				style2.setBorderTop(BorderStyle.THIN);
				style2.setTopBorderColor(IndexedColors.BLACK.getIndex());
				
				style2.setBorderLeft(BorderStyle.THIN);
				style2.setLeftBorderColor(IndexedColors.BLACK.getIndex());
				
				style2.setBorderBottom(BorderStyle.THIN);
				style2.setBottomBorderColor(IndexedColors.BLACK.getIndex());
			
				style2.setBorderRight(BorderStyle.THIN);
				style2.setRightBorderColor(IndexedColors.BLACK.getIndex());
				

				CellStyle style2f = master.createCellStyle();
				style2f.setDataFormat(createHelper.createDataFormat().getFormat("0.0%"));
				style2f.setBorderTop(BorderStyle.THIN);
				style2f.setTopBorderColor(IndexedColors.BLACK.getIndex());
				
				style2f.setBorderLeft(BorderStyle.THIN);
				style2f.setLeftBorderColor(IndexedColors.BLACK.getIndex());
				
				style2f.setBorderBottom(BorderStyle.THIN);
				style2f.setBottomBorderColor(IndexedColors.BLACK.getIndex());
			
				style2f.setBorderRight(BorderStyle.THIN);
				style2f.setRightBorderColor(IndexedColors.BLACK.getIndex());
				
				style2f.setFont(fred);
				
				CellStyle style2b = master.createCellStyle();
				style2b.setDataFormat(createHelper.createDataFormat().getFormat("0.0%"));
				style2b.setBorderTop(BorderStyle.THIN);
				style2b.setTopBorderColor(IndexedColors.BLACK.getIndex());
				
				style2b.setBorderLeft(BorderStyle.THIN);
				style2b.setLeftBorderColor(IndexedColors.BLACK.getIndex());
				
				style2b.setBorderBottom(BorderStyle.THIN);
				style2b.setBottomBorderColor(IndexedColors.BLACK.getIndex());
			
				style2b.setBorderRight(BorderStyle.THIN);
				style2b.setRightBorderColor(IndexedColors.BLACK.getIndex());
				style2b.setFont(font);
				
				//time [h]
				CellStyle style3 = master.createCellStyle();
				style3.setDataFormat(createHelper.createDataFormat().getFormat("[h]:mm"));
				style3.setBorderTop(BorderStyle.THIN);
				style3.setTopBorderColor(IndexedColors.BLACK.getIndex());
				
				style3.setBorderLeft(BorderStyle.THIN);
				style3.setLeftBorderColor(IndexedColors.BLACK.getIndex());
				
				style3.setBorderBottom(BorderStyle.THIN);
				style3.setBottomBorderColor(IndexedColors.BLACK.getIndex());
				
				style3.setBorderRight(BorderStyle.THIN);
				style3.setRightBorderColor(IndexedColors.BLACK.getIndex());
				
				CellStyle style3f = master.createCellStyle();
				style3f.setDataFormat(createHelper.createDataFormat().getFormat("[h]:mm"));
				style3f.setBorderTop(BorderStyle.THIN);
				style3f.setTopBorderColor(IndexedColors.BLACK.getIndex());
				
				style3f.setBorderLeft(BorderStyle.THIN);
				style3f.setLeftBorderColor(IndexedColors.BLACK.getIndex());
				
				style3f.setBorderBottom(BorderStyle.THIN);
				style3f.setBottomBorderColor(IndexedColors.BLACK.getIndex());
				
				style3f.setBorderRight(BorderStyle.THIN);
				style3f.setRightBorderColor(IndexedColors.BLACK.getIndex());
				
				style3f.setFont(fred);
				
				CellStyle style3b = master.createCellStyle();
				style3b.setDataFormat(createHelper.createDataFormat().getFormat("[h]:mm"));
				style3b.setBorderTop(BorderStyle.THIN);
				style3b.setTopBorderColor(IndexedColors.BLACK.getIndex());
				
				style3b.setBorderLeft(BorderStyle.THIN);
				style3b.setLeftBorderColor(IndexedColors.BLACK.getIndex());
				
				style3b.setBorderBottom(BorderStyle.THIN);
				style3b.setBottomBorderColor(IndexedColors.BLACK.getIndex());
				
				style3b.setBorderRight(BorderStyle.THIN);
				style3b.setRightBorderColor(IndexedColors.BLACK.getIndex());
				style3b.setFont(font);
			
			 //normal i.e. no format conditions
			 CellStyle style4 = master.createCellStyle();
				style4.setBorderTop(BorderStyle.THIN);
				style4.setTopBorderColor(IndexedColors.BLACK.getIndex());
				
				style4.setBorderLeft(BorderStyle.THIN);
				style4.setLeftBorderColor(IndexedColors.BLACK.getIndex());
				
				style4.setBorderBottom(BorderStyle.THIN);
				style4.setBottomBorderColor(IndexedColors.BLACK.getIndex());
				
				style4.setBorderRight(BorderStyle.THIN);
				style4.setRightBorderColor(IndexedColors.BLACK.getIndex());
				
				 CellStyle style4f = master.createCellStyle();
					style4f.setBorderTop(BorderStyle.THIN);
					style4f.setTopBorderColor(IndexedColors.BLACK.getIndex());
					
					style4f.setBorderLeft(BorderStyle.THIN);
					style4f.setLeftBorderColor(IndexedColors.BLACK.getIndex());
					
					style4f.setBorderBottom(BorderStyle.THIN);
					style4f.setBottomBorderColor(IndexedColors.BLACK.getIndex());
					
					style4f.setBorderRight(BorderStyle.THIN);
					style4f.setRightBorderColor(IndexedColors.BLACK.getIndex());
					
					style4f.setFont(fred);
				
				 String name = BMRep.list()[i].substring(0, BMRep.list()[i].indexOf("-"));
				 String region = BMRep.list()[i].substring(BMRep.list()[i].indexOf("-")+1);
				
				Sheet msheet = master.createSheet("New Sheet");
				
			Row mrow = msheet.createRow(0);
			Cell mcell = mrow.createCell(0);
			mcell.setCellStyle(style4);
			mcell.setCellValue("Name");
			mcell.setCellStyle(header);
			mcell =mrow.createCell(1);
			mcell.setCellValue(name);
			mcell.setCellStyle(style4);
			mrow = msheet.createRow(1);
			mcell = mrow.createCell(0);
			mcell.setCellValue("Region");
			mcell.setCellStyle(header);
			mcell = mrow.createCell(1);
			mcell.setCellValue(region);
			mcell.setCellStyle(style4);
			//reading data from targets file
			
			int top=0;
			
			
			//#1 BM TARGETS
			
			Workbook targets = WorkbookFactory.create(tfile);
		
			
			//target methods
			
			
			
			Sheet tsheet = targets.getSheetAt(0);
			
			Row trow = tsheet.getRow(2);
			Cell tcell = trow.getCell(0);
			
			mrow = msheet.createRow(3);
			mcell = mrow.createCell(0);
			mcell.setCellValue("Targets");
			mcell.setCellStyle(header);
			
			mrow = msheet.createRow(4);
			
			for (int k = 0; k < 5; k++) {
				
				tcell = trow.getCell(k);
				mcell = mrow.createCell(k);
				mcell.setCellStyle(header);
				mcell.setCellValue(tcell.getStringCellValue());
			}
			
			
			
			
			
			
			
			int start = 3;
			trow = tsheet.getRow(start++);
			tcell = trow.getCell(0);
			
			try {
				
				while(tcell.getStringCellValue().compareToIgnoreCase(name)!=0)
				{
					trow = tsheet.getRow(start++);
					tcell = trow.getCell(0);
				}
				
				
				//starting print for targets
				
				mrow = msheet.createRow(5);
				
				
				 top = mrow.getRowNum();	
				
				
				
				try {
					
					
					do
					{
						
						for (int k = 0; k < 5; k++) {
						
							tcell = trow.getCell(k);
							mcell = mrow.createCell(k);
							mcell.setCellStyle(style4);
							
							if(k==0)
							{
								mcell.setCellValue(tcell.getStringCellValue());
								if(mrow.getRowNum()==top)
									mcell.setCellStyle(header);
							}
							else
								if(k==4)
								{
									double num = tcell.getNumericCellValue();
									if(mrow.getRowNum()==top)
										mcell.setCellStyle(style2b);
									else
										mcell.setCellStyle(style2);
									mcell.setCellValue(num);
								}else
								{
									if(mrow.getRowNum()==top)
										mcell.setCellStyle(style1b);
									else
										if(k==3)
										{
											double hold = mrow.getCell(2).getNumericCellValue();
											if(tcell.getNumericCellValue()<hold)
												mcell.setCellStyle(style1f);
											else
												mcell.setCellStyle(style1);
										}
									else
										mcell.setCellStyle(style1);
									mcell.setCellValue(tcell.getNumericCellValue());
								}
						}
						
						trow = tsheet.getRow(trow.getRowNum()+1);
						tcell = trow.getCell(0);
						mrow = msheet.createRow(mrow.getRowNum()+1);
					}while(tcell.getStringCellValue().substring(tcell.getStringCellValue().length()-1).compareToIgnoreCase(")")==0);
					
					
					
				} catch (Exception e) {
					// 
				}
				
				
			}catch(Exception e)
			{
				System.out.println("no target informatoin on "+name);
				//continue;
			}
			
			
			
			
			
			
			
			
			
			
			//BEGINNING OF KRONOS #2
			
			//sheet with summary
			Sheet ksheet = kronos.getSheetAt(1);
			//start from row with heading titles
			Row krow = ksheet.getRow(1);
			Cell kcell = krow.getCell(0);
			
			mrow = msheet.createRow(mrow.getRowNum()+1);
			mcell = mrow.createCell(0);
			mcell.setCellValue("Kronos");
			mcell.setCellStyle(header);
				
			mrow = msheet.createRow(mrow.getRowNum()+1);
			mcell = mrow.createCell(0);
			mcell.setCellStyle(header);
			mcell.setCellValue("Client");
			
			
			for (int k = 1; k < 5; k++) {
				kcell = krow.getCell(k);
				mcell = mrow.createCell(k);
				mcell.setCellStyle(header);
				mcell.setCellValue(kcell.getStringCellValue());
			}
			double iv=0, av=0, ih=0, ah=0;
			try {
				
				krow = ksheet.getRow(1);
				kcell = krow.getCell(0);
				while(kcell.getStringCellValue().compareToIgnoreCase(name)!=0)
				{
					krow = ksheet.getRow(krow.getRowNum()+1);
					kcell = krow.getCell(0);
				}
				

				
				
				iv = krow.getCell(1).getNumericCellValue();
				av = krow.getCell(2).getNumericCellValue();
				ih = krow.getCell(3).getNumericCellValue();
				ah = krow.getCell(4).getNumericCellValue();
				
			}catch(Exception e)
			{
				System.out.println("no kronos informatoin on "+name);
				//continue;
			}
			
			try {
				
				mrow = msheet.createRow(mrow.getRowNum()+1);
				top = mrow.getRowNum();
				
				do
				{
					
					for (int k = 0; k < 5; k++) {
						
						mcell = mrow.createCell(k);
						
						if(mrow.getRowNum()==top)
							mcell.setCellStyle(header);
						else
							mcell.setCellStyle(style4);
						
						kcell = krow.getCell(k);		
						
						if(k==0)
						{
							mcell.setCellValue(kcell.getStringCellValue());
						}else
							if(k==2&&mrow.getRowNum()!=top)
							{
								if(kcell.getNumericCellValue()<mrow.getCell(1).getNumericCellValue())
									mcell.setCellStyle(style4f);
								
								mcell.setCellValue(kcell.getNumericCellValue());
							}
						else
							if(k>2)
							{
								if(mrow.getRowNum()==top)
									mcell.setCellStyle(style3b);
								else
									if(k==4)
									{
										if(kcell.getNumericCellValue()<mrow.getCell(3).getNumericCellValue())
											mcell.setCellStyle(style3f);
										else
											mcell.setCellStyle(style3);
									}else
										mcell.setCellStyle(style3);
								
								
								mcell.setCellValue(kcell.getNumericCellValue());
							}
							else
								
						{
							mcell.setCellValue(kcell.getNumericCellValue());
						}
						
						//END OF FOR LOOP HERE
						
					}
					
					
					/*INCREMENTORS*/
					mrow = msheet.createRow(mrow.getRowNum()+1);
					krow = ksheet.getRow(krow.getRowNum()+1);
					kcell = krow.getCell(0);
					
					
				}while(kcell.getStringCellValue().substring(kcell.getStringCellValue().length()-1).compareToIgnoreCase(")")==0);
				
				double avg = ((av/iv)+(ah/ih))/2;
				
				mcell = mrow.createCell(4);
				mcell.setCellValue(avg);
				mcell.setCellStyle(style2b);
				
			} catch (Exception e) {
				// 
			}
			
			for (int k = 0; k < 5; k++) {
				msheet.autoSizeColumn(k);
			}
			
			int arow = mrow.getRowNum()+1;
			
			
			
			
			
			
			//START OF DORMANT #3
			
			//specify where dormant will be on master sheet
			int mcol = 8;
			mrow = msheet.getRow(3);
			mcell = mrow.getCell(mcol);
			
			Sheet dsheet = dormant.getSheetAt(0);
			//TODO: changed 5 to 4 below
			Row drow = dsheet.getRow(4);
			Cell dcell = drow.getCell(0);
			
			
			int last = 0;
			
			//last is the number of columns
			//method iterates right until hits white space to get used columns
			try {
					do
					{
						dcell = drow.getCell(++last);
					}while(dcell.getStringCellValue().compareTo("")!=0);
					
					
			} catch (Exception e) {
				// TODO: handle exception
			}
			
			//come back to start after acquiring last
			drow = dsheet.getRow(4);
			
			//writting
			
			mrow = msheet.getRow(3);
			mcell = mrow.createCell(mcol);
		
			mcell.setCellStyle(header);
			
			try {
				//headers
				for (int j = 4; j <6 ; j++) {
					mcol = 8;
					drow = dsheet.getRow(j);
					
					for (int k = 0; k < 2; k++) {
					dcell = drow.getCell(k);
					mcell = mrow.createCell(mcol++);
					mcell.setCellStyle(header);
					if(j==4&&k==0)
						mcell.setCellValue("Dormant Stock");
					else
						mcell.setCellValue(dcell.getStringCellValue()); 
				}
					
				for (int j2 = last-6; j2 <last ; j2++) {
					dcell = drow.getCell(j2);
					mcell = mrow.createCell(mcol++);
					mcell.setCellStyle(header);
					mcell.setCellValue(dcell.getStringCellValue()); 
				
				}	
					
					mrow = msheet.getRow(mrow.getRowNum()+1);
					
				}
				
			} catch (Exception e) {
				// 
				//continue;
			}
			
			
			//End of headers writting
			
				
				drow = dsheet.getRow(drow.getRowNum()+1);
				
				
				int beg = drow.getRowNum()-1;
				drow = dsheet.getRow(beg);
				dcell = drow.getCell(0);
			
				//searches for information pertsining to respective rep
				try {
					do
				{
					
					
					
					if(dcell==null)
						continue;
					
					drow = dsheet.getRow(beg++);
					dcell = drow.getCell(0);
					
				}while(dcell.getStringCellValue().compareToIgnoreCase(name)!=0);
					
					
					dcell = drow.getCell(0);
					
					//Find how many records the rep has pertaining to them
					
					int end = 0;
					do
					{
						drow = dsheet.getRow(drow.getRowNum()+1);
						dcell = drow.getCell(0);
						end++;
						
						
					}while(dcell.getStringCellValue().compareTo("")==0);
				
					double achieved = 0;
					
					//this for loop is responsible for moving down table
					for (int j = beg-1; j <beg+end ; j++) {
						mcol = 8;
						drow = dsheet.getRow(j);
						
						for (int k = 0; k < 2; k++) 
						{
							dcell = drow.getCell(k);
							mcell = mrow.createCell(mcol++);
							mcell.setCellStyle(style4);
							mcell.setCellValue(dcell.getStringCellValue()); 
						}
						
						//this for loop is reponsible with populating fields in row
						last = last-6;
						
						double num1=0, num2=0;
						
						for (int j2 = 0; j2 < 2; j2++) 
						{
							
							
							for (int k = 0; k < 3; k++) 
							{
								dcell = drow.getCell(last++);
								
								if(j2==0&&k==0)
									num1 = dcell.getNumericCellValue();
								else
									if(j2==1&&k==0)
										num2 = dcell.getNumericCellValue();
								
								
								mcell = mrow.createCell(mcol++);
								if(k==0)
									mcell.setCellStyle(style1);
								else
									if(k==1)
										mcell.setCellStyle(style4);
									else
										if(k==2)
										{
											if(dcell.getNumericCellValue()<=0)
												mcell.setCellStyle(style2);
											else
												mcell.setCellStyle(style2f);
										}
								mcell.setCellValue(dcell.getNumericCellValue());
								
								
								
								
							}
							
							
						}
						
						
						mcell = mrow.createCell(mcol++);
						mcell.setCellStyle(style2b);
						
						//checks whether last row to place %
						if(j <beg+end-1)
						{
							if(num1==0&&num2==0)
							{
								mcell.setCellValue("Achieved");
								achieved++;
							}
						else
							if(num2-num1>0)
								{
									mcell.setCellValue("Failed");
									mcell.setCellStyle(style2f);
								}
							else
								{
									mcell.setCellValue("Achieved");
									achieved++;
								}
						
						}else
						{
							mcell.setCellValue(achieved/end);
						}
						
						
						
					/*for (int j2 = last-6; j2 <last ; j2++) {
						dcell = drow.getCell(j2);
						mcell = mrow.createCell(mcol++);
						mcell.setCellStyle(style4);
						mcell.setCellValue(dcell.getNumericCellValue()); 
					
					}	*/
						
						
						if(msheet.getRow(mrow.getRowNum()+1)==null)
							mrow = msheet.createRow(mrow.getRowNum()+1);
						else
							mrow = msheet.getRow(mrow.getRowNum()+1);
					}
				
					for (int j = 8; j < 16; j++) {
						msheet.autoSizeColumn(j);
						
					}
					
				} catch (Exception e) {
					// 
					
					System.out.println("no dormant information on "+BMRep.list()[i]);
					//continue;
				}
				
				
				

				Client client = Client.accessToken("0/4142a98407fa8119b20a20b59c8c95ae");

				
			
				
				client.headers.put("Asana-Disable", "string_ids,new_sections,new_user_task_lists");
				
			 	
				
				
				List<User> users = client.users.findAll().query("team", "733042431372914").execute();
				users.addAll(client.users.findAll().query("team", "1104216325044301").execute());
				
				
				
				//remove duplicate dean and claire
				for (int j = 0; j < users.size() ; j++) {
					if(users.get(j).gid.compareTo("510943895984379")==0)
					{
						users.remove(j);
						break;
					}
				}
				
				for (int j = 0; j < users.size() ; j++) {
					if(users.get(j).gid.compareTo("510943467751185")==0)
					{
						users.remove(j);
						break;
					}
					
				}
				
				//get tasks for user
				
				tDate[] arr = new tDate[5];
				
				
				arr[0] = new tDate("1234", new Date().getTime(), "today");
				
				//TODO: switch/enum for rep name to gid
				
				String ugid="";
			
			
			
			
				switch (rep.valueOf(name).ordinal()) {
				case 0: ugid = "1155943241735889";//Aidan
				break;
				case 1: ugid = "1143898803372993";//Fulu
				break;
				case 2: ugid = "662072092829294";//Hazel
				break;
				case 3: ugid = "662072092829271";//Jackie
				break;
				case 4: ugid = "1112174143264743";//Jolene
				break;
				case 5: ugid = "1163017665830673";//Karen
				break;
				case 6: ugid = "1119664339723870";//Kaylen
				break;
				case 7: ugid = "1138380100468293";//William
				break;
				case 8: ugid = "804048676446699";//Louis
				break;
				case 9: ugid = "1128742066746577";//Mohamed
				break;
				case 10: ugid = "1113516810972893";//Najbulo
				break;
				case 11: ugid = "1164595080703813";//Roberto
				break;
				case 12: ugid = "1152557295298594";//Roebeth
				break;
				case 13: ugid = "618873806134395";//Roland
				break;
				case 14: ugid = "1133859561713096";//Sipho
				break;
				case 15: ugid = "1108363137349882";//Thabo
				break;
				case 16: ugid = "1196557710846061";//KiYano
				break;
				case 17: ugid = "1200006066577647";//Lindiwe
				break;
				case 18: ugid = "1200728660300272";//Natasha
				break;
				case 19: ugid = "1201166032790665";//Jimmy
				break;
				case 20: ugid = "1201684458943755";//Aaron
				break;
				case 21: ugid = "1201670713789795";//Campbell
				break;
				case 22: ugid = "1201785607352422";//Lynette
				break;
				
				}
				
				//1201166032790665
				
				//assessing user tasks - completion rates
			//	for(int i =0; i< users.size();i++)
				//{
					
					List<Task> noDate = new ArrayList<Task>();
					
					
					double complete = 0, incomplete = 0;
					
					List<Task> tasks = client.tasks.findAll().query("assignee", ugid).query("workspace", "510943893270416").option("opt_fields", "completed_at, due_on").execute();
					
					for(int x = 0; x< tasks.size();x++)
					{
						Task t = client.tasks.findById(tasks.get(x).gid).execute(); 
						
						
						//System.out.println(t.name);
						
						
						
						try {
							
							//ignore if completed
							if(t.completed)
								complete++;
							else//if not completed and due date has passed
								if(!t.completed&&t.dueOn.getValue()<new Date().getTime())
							{
								incomplete++;
								
								//storing top 5 oldest
								for (int j = 0; j < arr.length; j++) {
									if(arr[j]==null)
									{
										arr[j] = new tDate(t.gid, t.dueOn.getValue(), t.name);
										break;
									}
									else
										//if due date is older than arr date
										if(t.dueOn.getValue()<arr[j].date.getTime())
										{
											if(j<4)
											{
												//shift values down from index j
												for (int j2 = 4; j2 > j; j2--) {
													
														arr[j2] = arr[j2-1]; 
												}
											}
											
											arr[j] = new tDate(t.gid, t.dueOn.getValue(), t.name);
											break;
											
										}
								}
								
							}
							
							
						} catch (Exception e) {
							// 
							
							boolean found=false;
							
							//check in noDate
							for (int j = 0; j < noDate.size(); j++) {
								if(noDate.get(j).gid.compareTo(t.gid)==0)
								{
									found = true;
									break;
								}
							}
							
							if(!found)
							{
								noDate.add(t);
								
								URL url = new URL("https://app.asana.com/api/1.0/tasks/"+t.gid+"?opt_fields=created_by");
								HttpURLConnection conn = (HttpURLConnection) url.openConnection();
								
								conn.setRequestProperty("Authorization", "Bearer 0/4142a98407fa8119b20a20b59c8c95ae");
								conn.setRequestProperty("Content_Type", "application/json");
								conn.setRequestMethod("GET");
								
								BufferedReader in = new BufferedReader(new InputStreamReader(conn.getInputStream()));
								String output;
								
								JsonParser parser = new JsonParser();
								
								StringBuffer response = new StringBuffer();
								while((output = in.readLine())!=null)
								{
									response.append(output);
								}
								
								in.close();
								System.out.println(response.toString());
								
								JsonElement jsonTree = parser.parse(response.toString());
								
								JsonObject jsonObject = jsonTree.getAsJsonObject();
								
								JsonElement data = jsonObject.get("data");
								
								JsonObject dataObj = data.getAsJsonObject();
								
								JsonElement created = dataObj.get("created_by");
								
								JsonObject createdObj = created.getAsJsonObject();
								
								JsonElement gid = createdObj.get("gid");
								
								String cgid = gid.getAsString();
								
								Date tempdate = new Date();
								
								
								tempdate = new Date(tempdate.getTime()+259200000);
								
								
						//		client.tasks.create().data("workspace", "510943893270416").data("assignee", cgid).data("name", t.name+": Please set Due Date/Get "+name+" to mark as completed").data("due_on", 1900+tempdate.getYear()+"-"+1+tempdate.getMonth()+"-"+tempdate.getDate()).execute();
								//System.out.println(t.name);
								
							}
							
						
						}
						
						
						
								
					}
					
					
					
					//System.out.println(name);
					if((mrow = msheet.getRow(arow))==null)
						mrow = msheet.createRow(arow);
					mcell = mrow.createCell(0);
					mcell.setCellValue("Asana");
					mcell.setCellStyle(header);
					
					
					if((msheet.getRow(mrow.getRowNum()+1))==null)
						mrow = msheet.createRow(mrow.getRowNum()+1);
					else
						mrow = msheet.getRow(mrow.getRowNum()+1);
					mcell = mrow.createCell(0);
					mcell.setCellValue("Rep");
					mcell.setCellStyle(header);
					mcell = mrow.createCell(1);
					mcell.setCellValue(name);
					mcell.setCellStyle(header);
					mcell = mrow.createCell(2);
					mcell.setCellValue("Oldest Incomplete tasks");
					mcell.setCellStyle(header);
					mcell = mrow.createCell(3);
					mcell.setCellValue("Due Date");
					mcell.setCellStyle(header);
					
					
					if((msheet.getRow(mrow.getRowNum()+1))==null)
						mrow = msheet.createRow(mrow.getRowNum()+1);
					else
						mrow = msheet.getRow(mrow.getRowNum()+1);
					mcell = mrow.createCell(0);
					mcell.setCellValue("% tasks completed");
					mcell.setCellStyle(header);
					mcell = mrow.createCell(1);
					mcell.setCellValue(Math.round((complete/(complete+incomplete))*10000.0)/100.0+"%");
					mcell.setCellStyle(style2b);
					
					mrow = msheet.getRow(mrow.getRowNum()-1);
					
					try {
						for (int j = 0; j < arr.length; j++) {
							if((msheet.getRow(mrow.getRowNum()+1))==null)
								mrow = msheet.createRow(mrow.getRowNum()+1);
							else
								mrow = msheet.getRow(mrow.getRowNum()+1);
							mcell = mrow.createCell(2);
							if(arr[j].desc.compareToIgnoreCase("today")==0)
								break;
							mcell.setCellValue(arr[j].desc);
							mcell.setCellStyle(style4);
							
							mcell = mrow.createCell(3);
							mcell.setCellValue(1900+arr[j].date.getYear()+"-"+(1+arr[j].date.getMonth())+"-"+arr[j].date.getDate());
							mcell.setCellStyle(style4);
						}
					} catch (Exception e) {
						// 
					}
					
					
					
			//START OF PHOTO INDEX
					
				/*	
					//uses store matrix to get full list of stores
					Workbook photos = WorkbookFactory.create(new File("G:\\My Drive\\Master Store, Vendor and article lists\\Dischem stores BM matrix master.xls"));
					Sheet psheet = photos.getSheetAt(0);
					int row = 0;
					int temp = 0;
					
					while(msheet.getRow(mrow.getRowNum()+1)!=null)
					{
						mrow = msheet.getRow(mrow.getRowNum()+1);
					}
					mrow = msheet.createRow(mrow.getRowNum()+2);
					mcell = mrow.createCell(0);
					mcell.setCellValue("Photo Index");
					mcell.setCellStyle(header);
					
					mrow = msheet.createRow(mrow.getRowNum()+1);
					mcell = mrow.createCell(0);
					mcell.setCellValue("Store");
					mcell.setCellStyle(header);
					
					//first row under "Stores"
					row = mrow.getRowNum()+1;
					
					mrow = msheet.createRow(row);
					
					Row prow = psheet.getRow(1);
					//store name
					Cell pcell = prow.getCell(1);
					
					
					//populates store column
					while((prow = psheet.getRow(prow.getRowNum()+1))!=null)
					{
						mcell = mrow.createCell(0);
						
						//rep name
						if(prow.getCell(3).getStringCellValue().compareToIgnoreCase(name)==0)
						{
							temp++;
							pcell = prow.getCell(1);
							mcell.setCellValue(pcell.getStringCellValue());
							mcell.setCellStyle(header);
							mrow = msheet.createRow(mrow.getRowNum()+1);
						}
						
					}

					
					
					
					for (int j = 0; j < VENDORS.list().length; j++) {
						if(VENDORS.list()[j].compareToIgnoreCase("desktop.ini") == 0)
							 continue;
						
						mrow = msheet.getRow(row-1);
						mcell = mrow.createCell(j+1);
						mcell.setCellStyle(header);
						//set vendor name in column header
						mcell.setCellValue(VENDORS.list()[j]);
						
						
							
							//loops through all stores looking for match
							for (int k2 = row; k2 < row+temp; k2++) {
								mrow = msheet.getRow(k2);
								try {
									File stores = new File("G:\\My Drive\\Sales Shared\\Principle Photo Upload\\"+BMRep.list()[i]+"\\Photos\\"+VENDORS.list()[j]+"\\"+msheet.getRow(k2).getCell(0).getStringCellValue());
								
									if(stores.list().length>1)
									{
										mcell = mrow.createCell(j+1);
										mcell.setCellValue("Y");
										mcell.setCellStyle(style4);
										
									}else
									{
										mcell = mrow.createCell(j+1);
										mcell.setCellValue("N");
										mcell.setCellStyle(style4);
									}
								
								} catch (Exception e) {
									// 
									
									mcell = mrow.createCell(j+1);
									mcell.setCellValue("-");
									mcell.setCellStyle(style4);
								}
								
								
							}
							
						
						
						
					}
					
					
					
					
				*/
			 
			 
			 //END OF KPI
			
			
				OutputStream out = new FileOutputStream(
						"G:\\My Drive\\Sales Shared\\4.Principle Photo Upload & KPI\\KPI\\" + BMRep.list()[i] + "\\"
								+ BMRep.list()[i] + " 30-"+month+"-2022.xls");
//				OutputStream out = new FileOutputStream(
//						"G:\\My Drive\\Sales Shared\\5.Principle Photo Upload & KPI\\KPI\\" + BMRep.list()[i] + "\\"
//								+ BMRep.list()[i] + " " + new Date().getDate() + "-"+ (new Date().getMonth()+1) +"-"+(new Date().getYear()-100)+".xls");
			master.write(out);
			master.close();
			System.out.println(i+1+"/"+(BMRep.list().length-1)+" completed "+name);
			 
		}
		 
		
			 System.out.println("Completed");
		
	
			
	}

}


