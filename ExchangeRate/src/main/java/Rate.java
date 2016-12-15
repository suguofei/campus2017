/**
 * Created by Wang on 2016/11/18.
 */
public class Rate {


    private String date;
    private int type;
    private double data;

    public Rate(String date, int type, double data){
        this.date=date;
        this.type=type;
        this.data=data;
    }

    public String  getDate(){
        return date;
    }
    public int getType(){
        return type;
    }
    public double getData(){
        return data;
    }

    @Override
    public String toString() {
        return date+"  "+type+"  "+data;
    }
}
