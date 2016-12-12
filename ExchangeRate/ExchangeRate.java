import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Scanner;

/**
 * Created by Wang on 2016/11/18.
 */
public class ExchangeRate {

    class MoneyType{
        public  static  final  int USD=0;
        public  static  final  int EUR=1;
        public  static  final  int HKD=2;
    }


    public  static  void main(String[] args){
        List<Rate> list=getRateList();
        Scanner scanner=new Scanner(System.in);
        System.out.println("请输入保存路径。");
        String filepath=scanner.nextLine();
        try {
            exportExcel(list,filepath);
            System.out.println("表格生成成功！");
        } catch (IOException e) {
            System.out.println("检查路径是否正确。");
            e.printStackTrace();
        }

    }


    public static List<Rate> getRateList(){
        List<Rate> rateList=new ArrayList<>();
        try {
            //解析源网页
            Document document=Jsoup.connect("http://www.chinamoney.com.cn/fe-c/historyParity.do").post();
            Element body=document.body();
            //获取汇率表格
            Element table=body.getElementsByTag("table").last();
            //获取所有行
            Elements rows=table.getElementsByTag("tr");
            for(int row=1;row<rows.size();row++ ){
                //获取所有列
                Elements tds=rows.get(row).getElementsByTag("td");
                String date=tds.get(0).getElementsByTag("div").get(0).text();
                rateList.add(new Rate(date, MoneyType.USD,Double.parseDouble(tds.get(1).text())));
                rateList.add(new Rate(date, MoneyType.EUR,Double.parseDouble(tds.get(2).text())));
                rateList.add(new Rate(date, MoneyType.HKD,Double.parseDouble(tds.get(4).text())));
            }
            System.out.println("源网页表格读取成功！");
        } catch (Exception e) {
            System.out.println("网络出现错误。");
            e.printStackTrace();
        }
        return rateList;
    }

    public  static  void exportExcel(List<Rate> list,String filepath) throws IOException {
        HSSFWorkbook wb = new HSSFWorkbook();
        HSSFSheet sheet=wb.createSheet("人民币汇率中间价");
        //创建表头
        HSSFRow row1=sheet.createRow(0);
        row1.createCell(0).setCellValue("日期");
        row1.createCell(1).setCellValue("CNY");
        row1.createCell(2).setCellValue("EUR");
        row1.createCell(3).setCellValue("HKD");

        for(int row=0;row<list.size();row+=3){
            HSSFRow row2=sheet.createRow(row+1);
            row2.createCell(0).setCellValue(list.get(row).getDate());
            row2.createCell(1).setCellValue(list.get(row).getData());
            row2.createCell(2).setCellValue(list.get(row+1).getData());
            row2.createCell(3).setCellValue(list.get(row+2).getData());

        }
        FileOutputStream out=new FileOutputStream(filepath);


        wb.write(out);
        out.close();

    }
}
