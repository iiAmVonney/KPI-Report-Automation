import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.time.LocalDateTime;
import java.util.Date;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.asana.Client;
import com.asana.models.Project;
import com.asana.models.Task;
import com.asana.models.User;

public class AsanaTasks {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub

		Client client = Client.accessToken("0/4142a98407fa8119b20a20b59c8c95ae");
		
		client.headers.put("Asana-Disable", "string_ids,new_sections");
		
		
		
		
		List<User> users = client.users.findAll().query("team", "733042431372914").execute();
		users.addAll(client.users.findAll().query("team", "1104216325044301").execute());
		
		System.out.println("hi");
		
		Workbook master =  new HSSFWorkbook();
		
		
		
		 CellStyle style4 = master.createCellStyle();
			style4.setBorderTop(BorderStyle.THIN);
			style4.setTopBorderColor(IndexedColors.BLACK.getIndex());
			
			style4.setBorderLeft(BorderStyle.THIN);
			style4.setLeftBorderColor(IndexedColors.BLACK.getIndex());
			
			style4.setBorderBottom(BorderStyle.THIN);
			style4.setBottomBorderColor(IndexedColors.BLACK.getIndex());
			
			style4.setBorderRight(BorderStyle.THIN);
			style4.setRightBorderColor(IndexedColors.BLACK.getIndex());
		
		
		
		for (int i = 0; i < users.size(); i++) {
			for (int j = i+1; j < users.size(); j++) {
				if(users.get(i).gid.compareTo(users.get(j).gid)==0)
				{
					users.remove(j);
					break;
				}
			}
			Sheet mSheeet = master.createSheet(users.get(i).name);
			
			
			Row mRow = mSheeet.createRow(0);
			Cell mCell = mRow.createCell(0);
			mCell.setCellStyle(style4);
			//current data and rep name
			mCell.setCellValue(LocalDateTime.now().toString());
			mCell = mRow.createCell(1);
			mCell.setCellValue(users.get(i).name);
			mCell.setCellStyle(style4);
			
			mRow = mSheeet.createRow(2);
			mCell = mRow.createCell(0);
			mCell.setCellValue("Task Name");
			mCell.setCellStyle(style4);
			
			mCell = mRow.createCell(1);
			mCell.setCellValue("Project");
			mCell.setCellStyle(style4);
			
			mCell = mRow.createCell(2);
			mCell.setCellValue("Due Date");
			mCell.setCellStyle(style4);
			
			List<Task> tasks = client.tasks.findAll().query("assignee", users.get(i).gid).query("workspace", "510943893270416").option("opt_fields", "completed_at, due_on").execute();
			
			for (int j = 0; j < tasks.size(); j++) {
				
				Task t = client.tasks.findById(tasks.get(j).gid).execute(); 
				
				
				
				if( t.completed)
					continue;
				
				
				
				mRow = mSheeet.createRow(mRow.getRowNum()+1);
				mCell = mRow.createCell(0);
				mCell.setCellValue(t.name);
				mCell.setCellStyle(style4);
				
				mCell =mRow.createCell(1);
				if(!t.projects.isEmpty())
				{
					Project[] arr = new Project[2];
					t.projects.toArray(arr);
					
					mCell.setCellValue(arr[0].name);
				}
				else
					mCell.setCellValue("Unassigned");
				mCell.setCellStyle(style4);
				
				mCell = mRow.createCell(2);
				if(t.dueOn!=null)
					mCell.setCellValue(t.dueOn.toString());
				else
					mCell.setCellValue("No Date Given");
				mCell.setCellStyle(style4);
				
					}
			mSheeet.autoSizeColumn(0);
			mSheeet.autoSizeColumn(1);
			mSheeet.autoSizeColumn(2);
			mSheeet.autoSizeColumn(3);
			
		}
		
		try {
			OutputStream out = new FileOutputStream("G:\\My Drive\\Sales Shared\\1.Reps Info\\9.Asana -Tasks Outstanding\\Asana "+(new Date().getYear()-100)+"-"+(new Date().getMonth()+1)+"-"+new Date().getDate()+".xls");
			master.write(out);
			System.out.println("On Drive");
		} catch (Exception e) {
			// TODO: handle exception
			OutputStream out = new FileOutputStream("C:\\Users\\Nick Von der Becke\\Desktop\\BlueMango\\Asana\\Asana "+new Date().getDate()+"-"+new Date().getMonth()+"-"+new Date().getYear()+".xls");
			master.write(out);
			System.out.println("On local host");
		}
		
		System.out.println("DONE");
	}

}
