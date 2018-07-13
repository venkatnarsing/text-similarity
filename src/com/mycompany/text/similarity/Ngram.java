package com.mycompany.text.similarity;

import java.util.ArrayList;
import java.util.List;

public class Ngram
{
    private class result
    {
        private String theWord;
        private int theCount;
 
        public result(String w, int c)
        {
            theWord = w;
            theCount = c;
        }
 
        public void setTheCount(int c)
        {
            theCount = c;
        }
 
        public String getTheWord()
        {
            return theWord;
        }
 
        public int getTheCount()
        {
            return theCount;
        }
    }
 
    private List<result> results;
 
    public Ngram()
    {
        results = new ArrayList<result>();
    }
    public Ngram(String str, int n)
    {
 
    }
 
    public double getSimilarity(String wordOne, String wordTwo, int n)
    {
        List<result> res1 = processString(wordOne, n);
        //displayResult(res1);
        List<result> res2 = processString(wordTwo, n);
        //displayResult(res2);
        int c = common(res1,res2);
        int u = union(res1,res2);
        double sim = (double)c/(double)u;
 
        return sim;
    }
 
    private int common(List<result> One, List<result> Two)
    {
        //int res = 0;
        List<String> commonTerms = new ArrayList<String>();
 
        for (int i = 0; i < One.size(); i++)
        {
            for (int j = 0; j < Two.size(); j++)
            {
                if (One.get(i).theWord.equalsIgnoreCase(Two.get(j).theWord)) {
                	//res++;
                	commonTerms.add(One.get(i).theWord);
                	break;
                }
            }
        }
 
        return commonTerms.size();
    }
 
    private int union(List<result> One, List<result> Two)
    {
        List<result> t = One;
 
        for (int i = 0; i < Two.size(); i++)
        {            
            boolean found = false;
            for (int j = 0; j < t.size() && !found; j++)
            {
                if (Two.get(i).theWord.equalsIgnoreCase(t.get(j).theWord))
                {
                    found = true;
                }
            }
 
            if (!found)
            {
                result r = Two.get(i);
                t.add(r);
            }
        }
 
        return t.size();
    }
 
    private List<result> processString(String c, int n)
    {
        List<result> t = new ArrayList<result>();
 
        //Adding single charactered term which decreasing the quality of final result
        /*String spacer = "";
        for (int i = 0; i < n-1; i++)
        {
            spacer = spacer + "%";
        }
        c = spacer + c + spacer;*/
         
        for (int i = 0; i < c.length(); i++)
        {
            if (i <= (c.length() - n))
            {
                if (contains(c.substring(i, n+i)) > 0)
                {
                    t.get(i).setTheCount(results.get(i).getTheCount()+1);
                }
                else
                {
                    t.add(new result(c.substring(i,n+i),1));
                }
            }
        }
        
        return t;
    }
 
    private int contains(String c)
    {
        for (int i = 0; i < results.size(); i++)
        {
            if (results.get(i).theWord.equalsIgnoreCase(c))
                return i;
        }
        return 0;
    }
 
    private void displayResult(List<result> d)
    {
        for (int i = 0; i < d.size(); i++)
        {
            System.out.println(d.get(i).theWord+" occurred "+d.get(i).theCount+" times");
        }
    }
    
    public static void main(String[] args) {
    	Ngram ng = new Ngram();
    	
    	double d = ng.getSimilarity("hero", "zero", 2);
    	System.out.println(d);
    }
}
