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
	// �ļ�����
	public static String fileVersion = "V1.3";
	public static String dirPath = "D:"; //
	public static String projName = "�������в�ҵ������ϵͳ"; //
	
	/*
	 * �������ݵ���excel
	 * 
	 * */
	public static void main(String[] args) {
		String result="";
		List listAll = new ArrayList();
		System.out.println("��ʼ��ȡ���б�ṹ����");
		try{
			List tbList = getTableList();
			System.out.println("��ȡ��");
			for(int i = 0;i<tbList.size();i++){
				String[] strings = (String[]) tbList.get(i);
				String tableName = strings[0].toString();
				List list = new ArrayList();
				list.add(tableName);
				list.add(getStructOfTable(tableName));
				System.out.println("�������ɱ�"+tableName+"�Ľṹ");
				listAll.add(list);
			}
			result = TableStructInfoToExcel(listAll,dirPath);
			System.out.println("�ѵ���");
		}catch(Exception e){
			e.printStackTrace();
			File file = new File(e.getMessage().toString());
			if(file.exists()){
				file.delete();
			}
		}
		System.out.println(result);
		
	}
	
	// ��ȡ���б���
	@SuppressWarnings("rawtypes")
	public static List getTableList(){
		String sql ="select TABNAME from syscat.TABLES where TABSCHEMA='AMARSCF' order by TABNAME";
		return getResult(sql,1);
	}
	// ��ȡ���б���-����
		@SuppressWarnings("rawtypes")
	public static List getTableListName(){
		String sql ="select REMARKS from syscat.TABLES where TABSCHEMA='AMARSCF' order by TABNAME";
		return getResult(sql,1);
	}
	// ��ȡ��Ľṹ
	@SuppressWarnings("rawtypes")
	public static List getStructOfTable(String tableName){
		String sql = "select COLNO+1,COLNAME,(case when (REMARKS is null) THEN COLNAME ELSE REMARKS END),TYPENAME,LENGTH,(case when (SCALE <>'0') THEN SCALE ELSE '' END),(case when (KEYSEQ is not null) THEN 'Y' END),NULLS,'' from syscat.columns where tabschema='AMARSCF' and tabname='"+tableName+"' order by COLNO";
		return getResult(sql,9);
	}
	// ����result list����
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
	// ���� Console ��ʾ
	
	// Table�ṹ ת�� excel 
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
		String[] tableField = {"���","�ֶ�Ӣ����","�ֶ�������","����","����","����","�Ƿ�����","�Ƿ�ɿ�","��ע"};
		int[] tableFieldWidth = {123,150,75,64,64,64,64,64,64};
		try{
			FileName = path + "\\" + projName +"_�����ֵ�"+fileVersion+".xls";//����·��
			fos = new FileOutputStream(FileName);
			HSSFWorkbook wb = new HSSFWorkbook(); //poi �д���������
			//HSSFSheet s = wb.createSheet();
			//wb.setSheetName(0, "���ݿ��ṹ");
			// Plain ����
			HSSFFont fontPlain = wb.createFont();
			fontPlain.setFontName("����");
			// ɾ��������
	        HSSFFont fontStrike = wb.createFont();
	        fontStrike.setStrikeout(true);//����ɾ����
	        fontStrike.setFontName("����");
	        // �Ӵ�Plain
			font = wb.createFont();
			font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);// ����Ӵ�
	        // ��ͷ����
	        HSSFFont fontTitle = wb.createFont();
	        fontTitle.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
	        fontTitle.setFontHeightInPoints((short)20);
	        fontTitle.setFontName("����");
	        // �����ͷ����
	        HSSFFont fontThead = wb.createFont();
	        fontThead.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);		// ����Ӵ�
	        fontThead.setFontHeightInPoints((short)11);				// �ֺ� 10.5
	        fontThead.setFontName("����");
	        // ������������
	        HSSFFont fontCol = wb.createFont();
	        fontCol.setFontHeightInPoints((short)11);				// �ֺ� 10.5
	        fontCol.setFontName("Times New Roman");
	        
	        // �Զ�����ɫ ����
	        HSSFPalette customPalette = wb.getCustomPalette();
	        customPalette.setColorAtIndex(HSSFColor.SKY_BLUE.index, (byte) 0, (byte) 176, (byte) 240); // ���
	        // ��ɫ��ͷ��ʽ
	        HSSFCellStyle BlueTitle = wb.createCellStyle();
			BlueTitle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
			BlueTitle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
			BlueTitle.setBorderTop(HSSFCellStyle.BORDER_THIN);
			BlueTitle.setBorderRight(HSSFCellStyle.BORDER_THIN);
			BlueTitle.setFont(fontThead);
			BlueTitle.setAlignment(HSSFCellStyle.ALIGN_CENTER);// ˮƽ����
			BlueTitle.setFillForegroundColor((short) HSSFColor.SKY_BLUE.index);//���ñ���ɫ - ����
			BlueTitle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
			// �����ֵ��ͷ��ʽ
			style = wb.createCellStyle();
			style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
			style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
			style.setBorderTop(HSSFCellStyle.BORDER_THIN);
			style.setBorderRight(HSSFCellStyle.BORDER_THIN);
			style.setFont(font);
			style.setAlignment(HSSFCellStyle.ALIGN_CENTER);// ˮƽ����
			style.setFillForegroundColor((short) HSSFColor.GREY_40_PERCENT.index);//���ñ���ɫ
			style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
			// ��ͨ������ʽ
			style2 = wb.createCellStyle();
			style2.setBorderBottom(HSSFCellStyle.BORDER_THIN);
			style2.setBorderLeft(HSSFCellStyle.BORDER_THIN);
			style2.setBorderTop(HSSFCellStyle.BORDER_THIN);
			style2.setBorderRight(HSSFCellStyle.BORDER_THIN);
			// ����������
	        HSSFFont linkFont= wb.createFont();  
	        linkFont.setUnderline((byte) 1);  
	        linkFont.setColor(HSSFColor.BLUE.index);  
			
			/* ��������ʽ*/  
	        HSSFCellStyle linkStyle = wb.createCellStyle();  
	        linkStyle.setFont(linkFont); 
	        /* ��������ʽ - �ӱ߿�*/  
	        HSSFCellStyle linkStyleB = wb.createCellStyle();  
	        linkStyleB.setFont(linkFont); 
	        linkStyleB.setBorderBottom(HSSFCellStyle.BORDER_THIN);
	        linkStyleB.setBorderLeft(HSSFCellStyle.BORDER_THIN);
	        linkStyleB.setBorderTop(HSSFCellStyle.BORDER_THIN);
	        linkStyleB.setBorderRight(HSSFCellStyle.BORDER_THIN);
	        // ���������ʽ
	        HSSFCellStyle titleStyle = wb.createCellStyle();
	        titleStyle.setFont(fontTitle);
	        titleStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);	// ˮƽ����
	        // ���������ʽ
	        HSSFCellStyle titleStyleL = wb.createCellStyle();
	        titleStyleL.setFont(fontTitle);
	        titleStyleL.setAlignment(HSSFCellStyle.ALIGN_LEFT);		// ˮƽ�����
	        
	        // ���� ��ͷ ��ʽ
	        HSSFCellStyle coverThead = wb.createCellStyle();
	        coverThead.setBorderBottom(HSSFCellStyle.BORDER_THIN);	// �߿�
	        coverThead.setBorderLeft(HSSFCellStyle.BORDER_THIN);
	        coverThead.setBorderTop(HSSFCellStyle.BORDER_THIN);
	        coverThead.setBorderRight(HSSFCellStyle.BORDER_THIN);
	        coverThead.setFont(fontThead);
	        coverThead.setAlignment(HSSFCellStyle.ALIGN_CENTER);	// ˮƽ����
	        // ���� ���� ��ʽ
 			HSSFCellStyle coverCol = wb.createCellStyle();
	        coverCol.setBorderBottom(HSSFCellStyle.BORDER_THIN);	// �߿�
	        coverCol.setBorderLeft(HSSFCellStyle.BORDER_THIN);
	        coverCol.setBorderTop(HSSFCellStyle.BORDER_THIN);
	        coverCol.setBorderRight(HSSFCellStyle.BORDER_THIN);
	        coverCol.setFont(fontCol);
	        coverCol.setAlignment(HSSFCellStyle.ALIGN_LEFT);		// ˮƽ�����
	        
	        

			// ��ȡ���б�����������б�
			List tbListName = getTableListName();
			
			// �����ֵ�ǰ������ҳ
			HSSFSheet s0 = wb.createSheet();
			s0.setDisplayGridlines(false);
			wb.setSheetName(0, "����");		//��һҳ
			row = s0.createRow(10); 		//��11��д����
			cell = row.createCell((short) 1);  
            cell.setCellValue("��ҵ������ϵͳ�����ֵ�");  
            cell.setCellStyle(titleStyle); 
            s0.addMergedRegion(new Region(10, (short) 1, 10,(short) 8));  // �ϲ�
            
            String[][] coverField = {
            		{"�ļ��汾��",fileVersion,"�ļ���ţ�",""},
            		{"�������ڣ�","","���ƣ�",""},
            		{"��    �ˣ�","","��׼��",""}
            };
            //coverField[1][1] = "2015/3/31";
            SimpleDateFormat currDate=new SimpleDateFormat("yyyy/MM/dd");
            coverField[1][1] = currDate.format(new Date());
            //��16�л����
            for(int i=0;i<3;i++){
            	int currLine = 15+i;
            	row = s0.createRow(currLine); 
            	for(int j=0;j<4;j++){
            		int jstart = j*2+1;
            		int jend = j*2+2;
            		cell = row.createCell((short) jstart);  
                    cell.setCellValue(coverField[i][j]);  
            		if(j%2==0){
            			cell.setCellStyle(coverThead); // �Ӵ�
            			cell = row.createCell((short) jstart+1);  
                        cell.setCellValue(""); 
                        cell.setCellStyle(coverThead); 
            		}else{
            			cell.setCellStyle(coverCol); // ���Ӵ�
            			cell = row.createCell((short) jstart+1);  
                        cell.setCellValue(""); 
                        cell.setCellStyle(coverCol); 
                    }
            		s0.addMergedRegion(new Region(currLine, (short) jstart, currLine,(short) jend));
            	}
            }
            
			HSSFSheet s1 = wb.createSheet();
			s1.setDisplayGridlines(false);
			wb.setSheetName(1, "�޶���¼");//�ڶ�ҳ
			row = s1.createRow(0);
			cell = row.createCell(1);
			cell.setCellValue("����ҳӦ����һ�汾�ֶ��������ǡ�");
			row = s1.createRow(3);
			cell = row.createCell(0);
			cell.setCellValue("�޶���¼");
			row = s1.createRow(4);
			cell = row.createCell(0);
			cell.setCellValue("�����г����ĵ����޶���¼�Ա�׷���ĵ��ı����ʷ���������Ķ���ÿ����ʽ�����汾���뱣���޶���¼��");
			
			String[] revTitle={"�޶��汾��","�޶���","�޶�����","�Ƿ������","�޶�����"};
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
			HSSFRichTextString richString = new HSSFRichTextString( "��ע����ɾ����Ϊ�ֶμ��٣�����Ϊ�����ֶ�" ); 
			richString.applyFont(fontPlain); 
			richString.applyFont( 4, 7, fontStrike );  
			cell.setCellValue(richString);
			
			HSSFSheet s2 = wb.createSheet();
			s2.setDisplayGridlines(false);
			wb.setSheetName(2, "���ݿ�Ŀ¼");//����ҳ
			row = s2.createRow(2); 		//��3��д����
			cell = row.createCell((short) 0);  
            cell.setCellValue("1. ���ݱ�һ��");  
            cell.setCellStyle(titleStyleL); 
            s2.addMergedRegion(new Region(2, (short) 0, 2,(short) 10));  // �ϲ�
            String[] indexCol={"���","���ݱ���������","���ݱ�Ӣ������","����˵��","ģʽ","����Ƶ��","������","��Ҫ","��¼������","��ע","�Ƿ񴴽�"};
            int[] indexWidth = {48,239,270,77,44,77,60,44,95,44,77};
            row = s2.createRow(4); 		//��3��д����
            for(int i=0;i<indexCol.length;i++){
            	s2.setColumnWidth((short) i,(short) indexWidth[i]*37);
            	cell = row.createCell(i);
            	cell.setCellValue(indexCol[i]);
            	cell.setCellStyle(BlueTitle);
            }
            
			// �����ֵ���ÿ�����Ӧһ��sheet
			for(int z=0;z<list.size();z++){
				// ��ȡӢ�ġ�������
				List listBean = (List)list.get(z);
				String[] names = (String[]) tbListName.get(z);
				System.out.println(listBean.get(0).toString());
				if("AWE_DO_JS_FUNCTION".equals(listBean.get(0).toString())){
					System.out.println("��������ʼ����");
				}
				// ��������һ����
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
				// �������� �½�������
				HSSFSheet s = wb.createSheet();
				//s.setDisplayGridlines(false);
				
				// ��һ�� ����
				row = s.createRow(currentRowNum);
				currentRowNum++;
				cell = row.createCell((short) 0);
				cell.setCellValue("");
				row.getCell((short) 0).setCellValue("");
				// - ��sheet λ���й� ��Ҫƫ�Ʋ���
				wb.setSheetName(z+3, listBean.get(0).toString());
				// �ڶ���  ��������	ҵ������¼��	����Ŀ¼
				row = s.createRow(currentRowNum);
				currentRowNum++;
				for(int i=0;i<2;i++){
					cell = row.createCell((short) i);
					cell.setCellValue("");
				}
				cell = row.createCell((short) 2);
				HSSFHyperlink hyperlink = new HSSFHyperlink(Hyperlink.LINK_DOCUMENT);  
	            
		        hyperlink.setAddress("#���ݿ�Ŀ¼!A1");  
		        cell.setHyperlink(hyperlink);  
		        // ���������ת  
		        cell.setCellValue("����Ŀ¼");  
		        cell.setCellStyle(linkStyle); 
				//cell.setCellType(HSSFCell.CELL_TYPE_FORMULA);
				//cell.setCellFormula("HYPERLINK(\"#���ݿ�Ŀ¼!A1\",\"����Ŀ¼\")");
			
				row.getCell((short) 0).setCellValue("��������");
				
				row.getCell((short) 1).setCellValue((names[0]==null)?listBean.get(0).toString():names[0].toString());
				// ������  ��Ӣ����	SCF_EXCEPTION_RECORD
				row = s.createRow(currentRowNum);
				currentRowNum++;
				for(int i=0;i<2;i++){
					cell = row.createCell((short) i);
					cell.setCellValue("");
				}
				row.getCell((short) 0).setCellValue("��Ӣ����");
				row.getCell((short) 1).setCellValue(listBean.get(0).toString());
				// ������  ��Ӣ����	SCF_EXCEPTION_RECORD
				row = s.createRow(currentRowNum);
				currentRowNum++;
				cell = row.createCell((short) 0);
				cell.setCellValue("");
				row.getCell((short) 0).setCellValue("���ñ�ģ��");
				/* �ϲ��ı�����
				row = s.createRow(currentRowNum);
				int pad = currentRowNum;
				currentRowNum++;
				for(int i=0;i<tableField.length;i++){
					cell = row.createCell((short) i);
					cell.setCellValue("");
					cell.setCellStyle(style);
				}
				row.getCell((short) 0).setCellValue("���ݿ��"+listBean.get(0).toString()+"�Ľṹ");
				*/
				// �����ڱ�ṹ����
				row = s.createRow(currentRowNum);
				currentRowNum++;
				for(int i=0;i<tableField.length;i++){
					//�������趨ÿһ�е�ֵ�Ϳ��
					cell = row.createCell((short) i);
					cell.setCellValue(new HSSFRichTextString(tableField[i]));
					cell.setCellStyle(style);
					s.setColumnWidth((short) i,(short) tableFieldWidth[i]*37);
				}
				
				//style.setFillForegroundColor((short) HSSFColor.WHITE.index);//���ñ���ɫ - ��
				
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
				//�ϲ���Ԫ��
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
