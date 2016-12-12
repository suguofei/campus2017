import java.io.*;
import java.util.regex.Pattern;

/**
 * Created by Wang on 2016/11/18.
 */
public class LinesCounter  {

    private  static  LinesCounter instance;
    public static  LinesCounter getInstance(){
        if(instance==null)
            instance=new LinesCounter();
        return instance;
    }


    public int getEffectiveLines(String path) throws FileNotFoundException,IOException {
        FileReader fileReader=new FileReader(path);
        BufferedReader bufferedReader=new BufferedReader(fileReader);
        String line="";
        int lines=0;
        while((line=bufferedReader.readLine())!=null){
            if(isEffectiveLine(line))
                lines++;
        }
        return  lines;
    }


    private boolean isEffectiveLine(String line){
        //单行注释匹配
        Pattern pattern=Pattern.compile("^\\s*//.*");
        //空行
        if(line.equals(""))
            return false;
        else if(pattern.matcher(line).matches())
            return false;
        else
            return true;
    }
}
