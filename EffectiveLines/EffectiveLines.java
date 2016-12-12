import java.io.IOException;
import java.util.Scanner;

/**
 * Created by Wang on 2016/11/18.
 */
public class EffectiveLines {


    public static  void main(String[] args){

        Scanner scanner=new Scanner(System.in);
        System.out.println("输入java文件路径,输入quit结束。");
        while (scanner.hasNext()) {
            String path=scanner.nextLine();
            if(path.equals("quit")){
                System.out.println("程序结束。");
                return;
            }
            try {
                int lines = LinesCounter.getInstance().getEffectiveLines(path);
                System.out.println(path + "\t有效行数：" + lines);

            } catch (IOException e) {
                System.out.println("统计失败，检查文件路径是否正确。");
                e.printStackTrace();
            }
        }
    }



}
