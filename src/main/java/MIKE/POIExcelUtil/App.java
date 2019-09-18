package MIKE.POIExcelUtil;

/**
 * Hello world!
 *
 */
public class App 
{
    public static void main( String[] args )
    {
    	ReportDefine reportDefine=new ReportDefine();
    	reportDefine.addSheet("第一个sheet");
    	String[] cells= {"111.12121","-12121","2019-09-09","zzhefdf","ceshiyiba"};
    	
    	reportDefine.addRow(cells,13);
    	reportDefine.setAutoWidthAll();
    	reportDefine.setColor(0, 0,23,11,false);
    	
    	
    	
    	reportDefine.addSheet("第二个sheet");
    	String[] cells2= {"1111111.1212","-232323","2019/09/19","zzhefdf","diergesheet"};
    	reportDefine.addRow(cells2,13);
    	reportDefine.setAutoWidthAll();
   
    	
    	reportDefine.addSheet("第三个sheet");
    	String[] mertitle= {"标题1","标题2","标题3","标题4","标题5"};
    	String[] title= {"标题1","标题2","标题3","标题4","标题5"};
    	String[] cells3= {"111111.1212","-232323","2019/09/19","zzhefdf","diergesheet"};
    	String[] cells4= {"111111.1212","-232323","2019/09/19","zzhefdf","diergesheet"};
    	String[] cells5= {"111111.1212","-232323","2019/09/19","zzhefdf","diergesheet"};
    	String[] cells6= {"111111.1212","-232323","2019/09/19","zzhefdf","diergesheet"};
    	reportDefine.addHeaderRow(mertitle,12);
    	reportDefine.addHeaderRow(title,12);
    	reportDefine.addRow(cells3,13);
    	reportDefine.addRow(cells4,13);
    	reportDefine.addRow(cells5,13);
    	reportDefine.addRow(cells6,13);
    	reportDefine.setAutoWidthAll();
    	reportDefine.setRowHight(0, 800);
    	
    	reportDefine.setColor(0, 0, 13,12,true);
    	reportDefine.mergeRow(0, 0, 0, 4);
    	reportDefine.mergeRow(1, 3, 1, 1);
    	
    	
    	reportDefine.addSheet("颜色代码表");
    	
    	for (int i = 0; i < 10; i++) {
    		reportDefine.createRow();
    		for (int j = 0; j <10; j++) {
    			String location=String.valueOf(i)+String.valueOf(j);
    			reportDefine.addCell(location);
			}	
		}
    	
    	for (int i = 0; i < 10; i++) {
    		for (int j = 0; j < 10; j++) {
    			int colorNum=0;
    			if (i==0) {
    				colorNum=j;
    			}else {
    				colorNum=Integer.valueOf(String.valueOf(i)+String.valueOf(j));
				}
    			reportDefine.setColor(i, j, colorNum,11,false);
			}
		}
    	
    	
    	reportDefine.addSheet("排序性能测试sheet");
    	for (int i = 0; i < 20000; i++) {
    		reportDefine.createRow();
    		for (int j = 0; j <100; j++) {
    			String location=String.valueOf(i)+String.valueOf(j);
    			reportDefine.addCell(location+"21ewdsewe2q1ewrwrwasdqewsadwarrewrsaf");
			}
		}
    	reportDefine.setAutoWidthAll();
    	
    	
    	String FileName="D:\\测试.xls";
    	reportDefine.outPutExcel(FileName);
    }
    
    
}
