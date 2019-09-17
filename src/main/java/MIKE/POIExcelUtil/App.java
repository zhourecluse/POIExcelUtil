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
    	
    	reportDefine.addRow(cells);
    	reportDefine.setAutoWidth();
    	
    	
    	
    	
    	reportDefine.addSheet("第二个sheet");
    	String[] cells2= {"1111111.1212","-232323","2019/09/19","zzhefdf","diergesheet"};
    	reportDefine.addRow(cells2);
    	reportDefine.setAutoWidth();
   
    	
    	reportDefine.addSheet("第三个sheet");
    	String[] cells3= {"111111.1212","-232323","2019/09/19","zzhefdf","diergesheet"};
    	reportDefine.addRow(cells3);
    	reportDefine.setAutoWidth();
    	
    	
    	
    	
    	String FileName="D:\\测试.xlsx";
    	
    	reportDefine.outPutExcel(FileName);
    }
    
    
}
