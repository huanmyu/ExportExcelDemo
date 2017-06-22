package com.bowen.controller;

import java.io.IOException;
import java.io.OutputStream;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.List;

import javax.servlet.ServletException;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

import com.bowen.bean.Student;

/**
 * Servlet implementation class ExcelServlet
 */
@WebServlet("/ExcelServlet")
public class ExcelServlet extends HttpServlet {
	static int d=2000;
	private static final long serialVersionUID = 1L;
       
    /**
     * @see HttpServlet#HttpServlet()
     */
    public ExcelServlet() {
        super();
        // TODO Auto-generated constructor stub
    }

	/**
	 * @see HttpServlet#doGet(HttpServletRequest request, HttpServletResponse response)
	 */
	protected void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		List<Student> list=new ArrayList<Student>();
		list.add(new Student(1,"aa",22,"男","金科路"));
		list.add(new Student(2,"bb",22,"女","大连路"));
		list.add(new Student(3,"cc",22,"男","张江高科"));
		list.add(new Student(4,"dd",22,"女","白杨林"));
		list.add(new Student(5,"ff",22,"男","芳华路"));
		list.add(new Student(6,"gg",22,"女","芳草路"));
		list.add(new Student(7,"hh",22,"男","芳芯路"));
		list.add(new Student(8,"ss",22,"女","浦电路"));
		// 创建一个Excel文件
		Workbook  wb=new HSSFWorkbook();
		// 创建一个Excel的Sheet
		Sheet sheet =wb.createSheet();
		// 设置列宽
		sheet.setColumnWidth(0, 6000);
		sheet.setColumnWidth(1, 6000);
		sheet.setColumnWidth(2, 6000);
		sheet.setColumnWidth(3, 6000);
		sheet.setColumnWidth(4, 6000);
		// 设置标头样式,大标题
		CellStyle  headerStyle1 =wb.createCellStyle();
		//设置水平居中
		headerStyle1.setAlignment(CellStyle.ALIGN_CENTER);
		Font headerFont1=wb.createFont();
		//设置字体样式
		headerFont1.setBoldweight(Font.BOLDWEIGHT_BOLD);
		headerFont1.setFontHeightInPoints((short) 18);
		headerStyle1.setFont(headerFont1);
		// 设置标头样式,小标题 创建标题样式1
		CellStyle headerStyle2 = wb.createCellStyle();
		headerStyle2.setAlignment(CellStyle.ALIGN_CENTER);
		// 设置子标题样式
		CellStyle headerStyle3 =wb.createCellStyle();
		headerStyle3.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		headerStyle3.setAlignment(CellStyle.ALIGN_CENTER);
		Font headerFont3 =wb.createFont();
		headerFont3.setBoldweight(Font.BOLDWEIGHT_BOLD);// 字体加粗
		headerFont3.setFontHeightInPoints((short) 10);
		headerStyle3.setFont(headerFont3);
		// 第一行大标题合并
		CellRangeAddress cra = new CellRangeAddress(0, 0, 0, 4);
		sheet.addMergedRegion(cra);
		// 第二行小标题合并
		CellRangeAddress cra1 = new CellRangeAddress(1, 1, 0, 4);
		sheet.addMergedRegion(cra1);
		// 第一行:大标题 单元格
		Row row1 = sheet.createRow(0);
		// 设置行高
		row1.setHeight((short) 600);
		Cell cell_1 = row1.createCell(0);
		//设置标题
		cell_1.setCellValue("学生信息统计报表");
		cell_1.setCellStyle(headerStyle1);
		// 第2行,小标题
		Row row2 = sheet.createRow(1);
		Cell cell_2 = row2.createCell(0);
		// 数据显示如下
		cell_2.setCellValue("2017年第13应届生信息");
		cell_2.setCellStyle(headerStyle2);
		// 第三行,子标题
		Row row3 = sheet.createRow(2);
		int totalCellIndex = 0;
		Cell cell3_1 = row3.createCell(totalCellIndex++);
		cell3_1.setCellValue("学号");
		cell3_1.setCellStyle(headerStyle3);
		Cell cell3_2 = row3.createCell(totalCellIndex++);
		cell3_2.setCellValue("姓名");
		cell3_2.setCellStyle(headerStyle3);
		Cell cell3_3 = row3.createCell(totalCellIndex++);
		cell3_3.setCellValue("年龄");
		cell3_3.setCellStyle(headerStyle3);
		Cell cell3_4 = row3.createCell(totalCellIndex++);
		cell3_4.setCellValue("性别");
		cell3_4.setCellStyle(headerStyle3);
		Cell cell3_5 = row3.createCell(totalCellIndex++);
		cell3_5.setCellValue("地址");
		cell3_5.setCellStyle(headerStyle3);
		//对单元格进行赋值操作
		if(list.size()>0){
			for(int i=0;i<list.size();i++){
				int j = 0;
				//从第4行开始给每一个单元格赋值（因为第一行是大标题，第二行是小标题，第三行是学生信息）
				Row row4 = sheet.createRow(i+3);
				//在第四行创建第一个单元格，用于存放学号
				Cell cell = row4.createCell(j++);
				cell.setCellValue(list.get(i).getId());
				cell.setCellStyle(headerStyle2);
				//在第五行创建第二个单元格，用于存放姓名
				Cell cell2 = row4.createCell(j++);
				cell2.setCellValue(list.get(i).getName());
				cell2.setCellStyle(headerStyle2);
				//在第五行创建第三个单元格，用于存放年龄
				Cell idcel3= row4.createCell(j++);
				idcel3.setCellValue(list.get(i).getAge());
				idcel3.setCellStyle(headerStyle2);
				//在第五行创建第四个单元格，用于存放性别
				Cell cell4 = row4.createCell(j++);
				cell4.setCellValue(list.get(i).getSex());
				cell4.setCellStyle(headerStyle2);
				//在第五行创建第五个单元格，用于存放地址
				Cell cell5 = row4.createCell(j++);
				cell5.setCellValue(list.get(i).getAddress());
				cell5.setCellStyle(headerStyle2);
			}
		}
		// 第六步，将文件存到指定位置
		String fileName = "StundetInfo";
		response.setContentType("application/vnd.ms-excel");
		response.setHeader("Content-disposition","attachment; filename=" + URLEncoder.encode(fileName + ".xls", "UTF-8"));
		OutputStream ouputStream = response.getOutputStream();
		wb.write(ouputStream);// 保存Excel文件
		ouputStream.flush();
		ouputStream.close();
		wb.close();
		response.flushBuffer();
	}

	/**
	 * @see HttpServlet#doPost(HttpServletRequest request, HttpServletResponse response)
	 */
	protected void doPost(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		// TODO Auto-generated method stub
		doGet(request, response);
	}

}
