package com.amarscf.app.lrsong.export;

import java.io.File;
import java.io.FileOutputStream;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFHyperlink;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.hssf.util.Region;
import org.apache.poi.ss.usermodel.Hyperlink;

public class DataToExcel {
	// 文件变量
	public static String fileVersion = "V1.3";
	public static String dirPath = "D:"; //
	public static String projName = "昆仑银行产业链金融系统"; //
	
	/*
	 * 按照数据导出excel
	 * 
	 * */
	public static void main(String[] args) {
		String result="";
		List listAll = new ArrayList();
		System.out.println("开始读取所有表结构数据");
		try{
			List tbList = getTableList();
			System.out.println("已取得");
			for(int i = 0;i<tbList.size();i++){
				String[] strings = (String[]) tbList.get(i);
				String tableName = strings[0].toString();
				List list = new ArrayList();
				list.add(tableName);
				list.add(getStructOfTable(tableName));
				System.out.println("正在生成表"+tableName+"的结构");
				listAll.add(list);
			}
			result = TableStructInfoToExcel(listAll,dirPath);
			System.out.println("已导入");
		}catch(Exception e){
			e.printStackTrace();
			File file = new File(e.getMessage().toString());
			if(file.exists()){
				file.delete();
			}
		}
		System.out.println(result);
		
	}
	
	// 获取所有表名
	@SuppressWarnings("rawtypes")
	public static List getTableList(){
		String sql ="select TABNAME from syscat.TABLES where TABSCHEMA='AMARSCF' order by TABNAME";
		return getResult(sql,1);
	}
	// 获取所有表名-中文
		@SuppressWarnings("rawtypes")
	public static List getTableListName(){
		String sql ="select REMARKS from syscat.TABLES where TABSCHEMA='AMARSCF' order by TABNAME";
		return getResult(sql,1);
	}
	// 获取表的结构
	@SuppressWarnings("rawtypes")
	public static List getStructOfTable(String tableName){
		String sql = "select COLNO+1,COLNAME,(case when (REMARKS is null) THEN COLNAME ELSE REMARKS END),TYPENAME,LENGTH,(case when (SCALE <>'0') THEN SCALE ELSE '' END),(case when (KEYSEQ is not null) THEN 'Y' END),NULLS,'' from syscat.columns where tabschema='AMARSCF' and tabname='"+tableName+"' order by COLNO";
		return getResult(sql,9);
	}
	// 返回result list集合
	public static List getResult(String sql,int length){
		List list = new ArrayList();
		ResultSet rs = null;
		ConnectionDb2 cd = new ConnectionDb2();
		try{
			rs = cd.executeQuery(sql);
			while(rs.next()){
				String[] string = new String[length];
				for(int i = 1;i<length+1;i++){
					string[i-1] = rs.getString(i);
				}
				list.add(string);
			}
			cd.close();
		}catch(SQLException e){
			e.printStackTrace();
		}
		return list;
	}
	// 调试 Console 显示
	
	// Table结构 转到 excel 
	@SuppressWarnings({ "deprecation", "rawtypes" })
	public static String TableStructInfoToExcel(List list,String path) throws Exception{
		String FileName="";
		FileOutputStream fos = null;
		HSSFRow row = null;
		HSSFCell cell = null;
		HSSFCellStyle style = null;
		HSSFCellStyle style2 = null;
		HSSFFont font = null;
		// int currentRowNum = 0;
		String[] tableField = {"序号","字段英文名","字段中文名","类型","长度","精度","是否主键","是否可空","备注"};
		int[] tableFieldWidth = {123,150,75,64,64,64,64,64,64};
		try{
			FileName = path + "\\" + projName +"_数据字典"+fileVersion+".xls";//生成路径
			fos = new FileOutputStream(FileName);
			HSSFWorkbook wb = new HSSFWorkbook(); //poi 中创建工作簿
			//HSSFSheet s = wb.createSheet();
			//wb.setSheetName(0, "数据库表结构");
			// Plain 字体
			HSSFFont fontPlain = wb.createFont();
			fontPlain.setFontName("宋体");
			// 删除线字体
	        HSSFFont fontStrike = wb.createFont();
	        fontStrike.setStrikeout(true);//设置删除线
	        fontStrike.setFontName("宋体");
	        // 加粗Plain
			font = wb.createFont();
			font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);// 字体加粗
	        // 标头字体
	        HSSFFont fontTitle = wb.createFont();
	        fontTitle.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
	        fontTitle.setFontHeightInPoints((short)20);
	        fontTitle.setFontName("宋体");
	        // 封面标头字体
	        HSSFFont fontThead = wb.createFont();
	        fontThead.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);		// 字体加粗
	        fontThead.setFontHeightInPoints((short)11);				// 字号 10.5
	        fontThead.setFontName("宋体");
	        // 封面正文字体
	        HSSFFont fontCol = wb.createFont();
	        fontCol.setFontHeightInPoints((short)11);				// 字号 10.5
	        fontCol.setFontName("Times New Roman");
	        
	        // 自定义颜色 靛蓝
	        HSSFPalette customPalette = wb.getCustomPalette();
	        customPalette.setColorAtIndex(HSSFColor.SKY_BLUE.index, (byte) 0, (byte) 176, (byte) 240); // 编号
	        // 蓝色表头样式
	        HSSFCellStyle BlueTitle = wb.createCellStyle();
			BlueTitle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
			BlueTitle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
			BlueTitle.setBorderTop(HSSFCellStyle.BORDER_THIN);
			BlueTitle.setBorderRight(HSSFCellStyle.BORDER_THIN);
			BlueTitle.setFont(fontThead);
			BlueTitle.setAlignment(HSSFCellStyle.ALIGN_CENTER);// 水平居中
			BlueTitle.setFillForegroundColor((short) HSSFColor.SKY_BLUE.index);//设置背景色 - 淡蓝
			BlueTitle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
			// 数据字典表头样式
			style = wb.createCellStyle();
			style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
			style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
			style.setBorderTop(HSSFCellStyle.BORDER_THIN);
			style.setBorderRight(HSSFCellStyle.BORDER_THIN);
			style.setFont(font);
			style.setAlignment(HSSFCellStyle.ALIGN_CENTER);// 水平居中
			style.setFillForegroundColor((short) HSSFColor.GREY_40_PERCENT.index);//设置背景色
			style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
			// 普通字体样式
			style2 = wb.createCellStyle();
			style2.setBorderBottom(HSSFCellStyle.BORDER_THIN);
			style2.setBorderLeft(HSSFCellStyle.BORDER_THIN);
			style2.setBorderTop(HSSFCellStyle.BORDER_THIN);
			style2.setBorderRight(HSSFCellStyle.BORDER_THIN);
			// 超链接字体
	        HSSFFont linkFont= wb.createFont();  
	        linkFont.setUnderline((byte) 1);  
	        linkFont.setColor(HSSFColor.BLUE.index);  
			
			/* 超链接样式*/  
	        HSSFCellStyle linkStyle = wb.createCellStyle();  
	        linkStyle.setFont(linkFont); 
	        /* 超链接样式 - 加边框*/  
	        HSSFCellStyle linkStyleB = wb.createCellStyle();  
	        linkStyleB.setFont(linkFont); 
	        linkStyleB.setBorderBottom(HSSFCellStyle.BORDER_THIN);
	        linkStyleB.setBorderLeft(HSSFCellStyle.BORDER_THIN);
	        linkStyleB.setBorderTop(HSSFCellStyle.BORDER_THIN);
	        linkStyleB.setBorderRight(HSSFCellStyle.BORDER_THIN);
	        // 封面标题样式
	        HSSFCellStyle titleStyle = wb.createCellStyle();
	        titleStyle.setFont(fontTitle);
	        titleStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);	// 水平居中
	        // 封面标题样式
	        HSSFCellStyle titleStyleL = wb.createCellStyle();
	        titleStyleL.setFont(fontTitle);
	        titleStyleL.setAlignment(HSSFCellStyle.ALIGN_LEFT);		// 水平左对齐
	        
	        // 封面 标头 样式
	        HSSFCellStyle coverThead = wb.createCellStyle();
	        coverThead.setBorderBottom(HSSFCellStyle.BORDER_THIN);	// 边框
	        coverThead.setBorderLeft(HSSFCellStyle.BORDER_THIN);
	        coverThead.setBorderTop(HSSFCellStyle.BORDER_THIN);
	        coverThead.setBorderRight(HSSFCellStyle.BORDER_THIN);
	        coverThead.setFont(fontThead);
	        coverThead.setAlignment(HSSFCellStyle.ALIGN_CENTER);	// 水平居中
	        // 封面 正文 样式
 			HSSFCellStyle coverCol = wb.createCellStyle();
	        coverCol.setBorderBottom(HSSFCellStyle.BORDER_THIN);	// 边框
	        coverCol.setBorderLeft(HSSFCellStyle.BORDER_THIN);
	        coverCol.setBorderTop(HSSFCellStyle.BORDER_THIN);
	        coverCol.setBorderRight(HSSFCellStyle.BORDER_THIN);
	        coverCol.setFont(fontCol);
	        coverCol.setAlignment(HSSFCellStyle.ALIGN_LEFT);		// 水平左对齐
	        
	        

			// 获取所有表的中文名称列表
			List tbListName = getTableListName();
			
			// 数据字典前边若干页
			HSSFSheet s0 = wb.createSheet();
			s0.setDisplayGridlines(false);
			wb.setSheetName(0, "封面");		//第一页
			row = s0.createRow(10); 		//第11行写标题
			cell = row.createCell((short) 1);  
            cell.setCellValue("产业链金融系统数据字典");  
            cell.setCellStyle(titleStyle); 
            s0.addMergedRegion(new Region(10, (short) 1, 10,(short) 8));  // 合并
            
            String[][] coverField = {
            		{"文件版本：",fileVersion,"文件编号：",""},
            		{"发布日期：","","编制：",""},
            		{"审    核：","","批准：",""}
            };
            //coverField[1][1] = "2015/3/31";
            SimpleDateFormat currDate=new SimpleDateFormat("yyyy/MM/dd");
            coverField[1][1] = currDate.format(new Date());
            //第16行画表格
            for(int i=0;i<3;i++){
            	int currLine = 15+i;
            	row = s0.createRow(currLine); 
            	for(int j=0;j<4;j++){
            		int jstart = j*2+1;
            		int jend = j*2+2;
            		cell = row.createCell((short) jstart);  
                    cell.setCellValue(coverField[i][j]);  
            		if(j%2==0){
            			cell.setCellStyle(coverThead); // 加粗
            			cell = row.createCell((short) jstart+1);  
                        cell.setCellValue(""); 
                        cell.setCellStyle(coverThead); 
            		}else{
            			cell.setCellStyle(coverCol); // 不加粗
            			cell = row.createCell((short) jstart+1);  
                        cell.setCellValue(""); 
                        cell.setCellStyle(coverCol); 
                    }
            		s0.addMergedRegion(new Region(currLine, (short) jstart, currLine,(short) jend));
            	}
            }
            
			HSSFSheet s1 = wb.createSheet();
			s1.setDisplayGridlines(false);
			wb.setSheetName(1, "修订记录");//第二页
			row = s1.createRow(0);
			cell = row.createCell(1);
			cell.setCellValue("【本页应从上一版本手动拷贝覆盖】");
			row = s1.createRow(3);
			cell = row.createCell(0);
			cell.setCellValue("修订记录");
			row = s1.createRow(4);
			cell = row.createCell(0);
			cell.setCellValue("本节列出该文档的修订记录以便追踪文档的变更历史，并方便阅读。每个正式发布版本必须保留修订记录。");
			
			String[] revTitle={"修订版本号","修订人","修订日期","是否已添加","修订描述"};
			int[] revWidth = {92,74,82,95,413};
			
			for(int j=0;j<7;j++){
				row = s1.createRow(5+j);
				for(int i=0;i<5;i++){
					cell = row.createCell(i);
					if(j==0){
						s1.setColumnWidth((short) i,(short) revWidth[i]*37);
						cell.setCellValue(revTitle[i]);
						cell.setCellStyle(style); 
					}else{
						cell.setCellStyle(style2); 
					}
				}
			}
			row = s1.createRow(12);
			cell = row.createCell(0);
			HSSFRichTextString richString = new HSSFRichTextString( "备注：加删除线为字段减少，红字为新增字段" ); 
			richString.applyFont(fontPlain); 
			richString.applyFont( 4, 7, fontStrike );  
			cell.setCellValue(richString);
			
			HSSFSheet s2 = wb.createSheet();
			s2.setDisplayGridlines(false);
			wb.setSheetName(2, "数据库目录");//第三页
			row = s2.createRow(2); 		//第3行写标题
			cell = row.createCell((short) 0);  
            cell.setCellValue("1. 数据表一览");  
            cell.setCellStyle(titleStyleL); 
            s2.addMergedRegion(new Region(2, (short) 0, 2,(short) 10));  // 合并
            String[] indexCol={"序号","数据表中文名称","数据表英文名称","特殊说明","模式","更新频率","更新者","概要","记录保存期","备注","是否创建"};
            int[] indexWidth = {48,239,270,77,44,77,60,44,95,44,77};
            row = s2.createRow(4); 		//第3行写标题
            for(int i=0;i<indexCol.length;i++){
            	s2.setColumnWidth((short) i,(short) indexWidth[i]*37);
            	cell = row.createCell(i);
            	cell.setCellValue(indexCol[i]);
            	cell.setCellStyle(BlueTitle);
            }
            
			// 数据字典中每个表对应一个sheet
			for(int z=0;z<list.size();z++){
				// 获取英文、中文名
				List listBean = (List)list.get(z);
				String[] names = (String[]) tbListName.get(z);
				System.out.println(listBean.get(0).toString());
				if("AWE_DO_JS_FUNCTION".equals(listBean.get(0).toString())){
					System.out.println("接下来开始出错");
				}
				// 逐条插入一览表
				row=s2.createRow(z+5);
				for(int i=0;i<indexCol.length;i++){
					cell = row.createCell(i);
					switch(i){
						case 0:
							cell.setCellValue(z+1);
							cell.setCellStyle(style2);
							break;
						case 1:
							if(names[0]==null){
								cell.setCellValue(listBean.get(0).toString());
							}else{
								cell.setCellValue(names[0].toString());
							}
							cell.setCellStyle(style2);
							break;
						case 2:
							Hyperlink hyperlink = new HSSFHyperlink(Hyperlink.LINK_DOCUMENT); 
					        hyperlink.setAddress("#"+listBean.get(0).toString()+"!A1");  
					        cell.setHyperlink(hyperlink);  
					        cell.setCellValue(listBean.get(0).toString());  
					        cell.setCellStyle(linkStyleB); 
							break;
						default:
							cell.setCellValue("");
							cell.setCellStyle(style2);
					}	
				}
				int currentRowNum = 0;
				// 工作簿内 新建工作表
				HSSFSheet s = wb.createSheet();
				//s.setDisplayGridlines(false);
				
				// 第一行 空行
				row = s.createRow(currentRowNum);
				currentRowNum++;
				cell = row.createCell((short) 0);
				cell.setCellValue("");
				row.getCell((short) 0).setCellValue("");
				// - 与sheet 位置有关 需要偏移操作
				wb.setSheetName(z+3, listBean.get(0).toString());
				// 第二行  表中文名	业务变更记录表	返回目录
				row = s.createRow(currentRowNum);
				currentRowNum++;
				for(int i=0;i<2;i++){
					cell = row.createCell((short) i);
					cell.setCellValue("");
				}
				cell = row.createCell((short) 2);
				HSSFHyperlink hyperlink = new HSSFHyperlink(Hyperlink.LINK_DOCUMENT);  
	            
		        hyperlink.setAddress("#数据库目录!A1");  
		        cell.setHyperlink(hyperlink);  
		        // 点击进行跳转  
		        cell.setCellValue("返回目录");  
		        cell.setCellStyle(linkStyle); 
				//cell.setCellType(HSSFCell.CELL_TYPE_FORMULA);
				//cell.setCellFormula("HYPERLINK(\"#数据库目录!A1\",\"返回目录\")");
			
				row.getCell((short) 0).setCellValue("表中文名");
				
				row.getCell((short) 1).setCellValue((names[0]==null)?listBean.get(0).toString():names[0].toString());
				// 第三行  表英文名	SCF_EXCEPTION_RECORD
				row = s.createRow(currentRowNum);
				currentRowNum++;
				for(int i=0;i<2;i++){
					cell = row.createCell((short) i);
					cell.setCellValue("");
				}
				row.getCell((short) 0).setCellValue("表英文名");
				row.getCell((short) 1).setCellValue(listBean.get(0).toString());
				// 第四行  表英文名	SCF_EXCEPTION_RECORD
				row = s.createRow(currentRowNum);
				currentRowNum++;
				cell = row.createCell((short) 0);
				cell.setCellValue("");
				row.getCell((short) 0).setCellValue("引用表模板");
				/* 合并的标题行
				row = s.createRow(currentRowNum);
				int pad = currentRowNum;
				currentRowNum++;
				for(int i=0;i<tableField.length;i++){
					cell = row.createCell((short) i);
					cell.setCellValue("");
					cell.setCellStyle(style);
				}
				row.getCell((short) 0).setCellValue("数据库表"+listBean.get(0).toString()+"的结构");
				*/
				// 创建第表结构内容
				row = s.createRow(currentRowNum);
				currentRowNum++;
				for(int i=0;i<tableField.length;i++){
					//创建并设定每一列的值和宽度
					cell = row.createCell((short) i);
					cell.setCellValue(new HSSFRichTextString(tableField[i]));
					cell.setCellStyle(style);
					s.setColumnWidth((short) i,(short) tableFieldWidth[i]*37);
				}
				
				//style.setFillForegroundColor((short) HSSFColor.WHITE.index);//设置背景色 - 白
				
				List list2 = (List) listBean.get(1);
				for(int i=0;i < list2.size();i++){
					row = s.createRow(currentRowNum);
					currentRowNum++;
					String[] strings = (String[]) list2.get(i);
					for(int j=0;j<strings.length;j++){
						cell = row.createCell((short)j);
						cell.setCellValue(new HSSFRichTextString(strings[j]));
						cell.setCellStyle(style2);
					}
				}
				//合并单元格
				//s.addMergedRegion(new Region(pad,(short) 0,pad,(short)(tableField.length - 1)));
				currentRowNum++;
			}
			wb.write(fos);
			fos.close();
		}catch(Exception e){
			e.printStackTrace();
			fos.close();
			throw new Exception(FileName);
		}
		
		return FileName;
		
	}
}
