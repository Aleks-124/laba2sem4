package org.example;
import java.util.HashMap;
import org.apache.commons.math3.stat.StatUtils;


public class ArithmeticMean {
    private  double[] dataarray;
    private HashMap<String, Double> dictionary;
    public ArithmeticMean(HashMap<String, Double> dictionary,double[] dataarray){
        this.dictionary=dictionary;
        this.dataarray=dataarray;
    }
    public  HashMap<String, Double> getArithmeticMean(HashMap<String, Double> dictionary,double[] dataarray) {
        dictionary.put("Среднее арифметическое", StatUtils.mean(dataarray));
        return dictionary;
    }
}