import com.spire.xls.Workbook;

public class ExcelToPDF {
	 public static void main(String[] args) {

	        System.out.println("---------开始----------");
	        //创建一个Workbook实例并加载Excel文件
	        Workbook workbook = new Workbook();
	        workbook.loadFromFile("C:\\Users\\Administrator\\Desktop\\2.xlsx");

	        //设置转换后PDF的页面宽度适应工作表的内容宽度
	        workbook.getConverterSetting().setSheetFitToWidth(true);

	        //转换为PDF并将生成的文档保存到指定路径
	        workbook.saveToFile("C:\\Users\\Administrator\\Desktop\\2.pdf");
	        System.out.println("---------成功----------");
	    }
}
