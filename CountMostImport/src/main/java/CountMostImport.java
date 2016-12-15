import java.util.Iterator;
import java.util.Map;
import java.util.Scanner;

public class CountMostImport {


    public  static  void main(String[] args){

       Scanner scanner=new Scanner(System.in);
        System.out.print("请输入源文件夹路径：");
        String path=scanner.nextLine();
       // String path="C:\\Program Files\\Java\\jdk1.8.0_20\\src";
        Map<String,Integer> result;
        result= CountImport.getInstance()
                .init()
                .countImport(path)
                .sort()
                .getResult();

        System.out.println("统计结果如下：");
        Iterator it = result.entrySet().iterator();
        int count=1;
        while(it.hasNext()){
            Map.Entry entry = (Map.Entry) it.next();
            System.out.println(count+":"+entry.getKey() + ":" + entry.getValue());
            count++;
            if(count==11)
                break;
        }
    }


}
