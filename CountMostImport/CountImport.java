import java.io.*;
import java.util.*;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Created by Wang on 2016/11/19.
 */
public class CountImport {

    private Map<String,Integer> counts;

    private static CountImport instance;

    public static CountImport getInstance(){
        if(instance==null)
            instance=new CountImport();
        return instance;
    }

    public CountImport init(){
        counts=new HashMap<>();
        return  this;
    }

    //返回排序后的结果
    public CountImport sort(){
        ArrayList<Map.Entry<String, Integer>> list = new ArrayList<>(counts.entrySet());
        Collections.sort(list,
                                        (o1, o2) ->o2.getValue() - o1.getValue());
        counts =new LinkedHashMap<>();
        for (int i = 0; i < list.size(); i++) {
            counts.put(list.get(i).getKey(), list.get(i).getValue());
        }
        return  this;
    }

    public Map<String,Integer> getResult(){
        return counts;
    }

    public CountImport countImport(String path){
        File file=new File(path);
        if(file.isDirectory()){
            String[] fileList=file.list();
            for(String str:fileList) {
                //如果不是java文件
                countImport(path + "\\" + str);
            }
        }else{
            //是java文件
            if(path.substring(path.lastIndexOf('.')+1).equals("java"))
                countImportOfFile(path,counts);
            return this;
        }
        return this;
    }

    private void countImportOfFile(String path,Map<String,Integer> map){

        try {
            FileReader fileReader=new FileReader(path);
            BufferedReader bufferedReader=new BufferedReader(fileReader);
            String line="";
            Pattern pattern=Pattern.compile("\\s*import\\s+[\\w\\.\\*]*");
            while((line=bufferedReader.readLine())!=null){
                Matcher matcher=pattern.matcher(line);
                //匹配成功
                if(matcher.find()){
                    String className=matcher.group().replaceAll("\\s*","").substring(6);
                   //除去静态导入成员对象的情况
                    if(className.equals("static"))
                        continue;

                    if(map.containsKey(className)){
                        map.put(className,map.get(className)+1);
                    }else {
                        map.put(className,1);
                    }
                }
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

    }



}
