package com.mycompany.text.similiraty;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.DecimalFormat;
import java.util.Iterator;

import org.apache.commons.codec.language.DoubleMetaphone;
import org.apache.commons.codec.language.RefinedSoundex;
import org.apache.commons.codec.language.Soundex;
import org.apache.commons.text.similarity.JaroWinklerDistance;
import org.apache.commons.text.similarity.LevenshteinDetailedDistance;
import org.apache.commons.text.similarity.LevenshteinResults;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class TextMatch {
	
	private static final String SRC_FILE_NAME = "Input_File.xlsx";
	private static final String DIST_FILE_NAME = "Results.xlsx";
	
	public static void main(String[] args) throws Exception {		
		readExcel();
	}
	
	private static void readExcel() throws Exception {
		FileInputStream excelFile = new FileInputStream(new File(SRC_FILE_NAME));
        Workbook workbook = new XSSFWorkbook(excelFile);
        Sheet datatypeSheet = workbook.getSheetAt(0);
        Iterator<Row> iterator = datatypeSheet.iterator();
        
        LevenshteinDetailedDistance ldd = new LevenshteinDetailedDistance();
        
        JaroWinklerDistance jwd = new JaroWinklerDistance();
        
        RefinedSoundex fs = new RefinedSoundex();
        
        Soundex soundex = new Soundex();
        
        DoubleMetaphone dm = new DoubleMetaphone();
        
        Ngram ng = new Ngram();
        
        DecimalFormat df = new DecimalFormat("#.##");
        
        //skip the header row.
        if(iterator.hasNext()){
        	iterator.next();
        }

        while (iterator.hasNext()) {

            Row currentRow = iterator.next();
            
            String primeBusinessName = currentRow.getCell(0).getStringCellValue();
            String whoisBusinessName = currentRow.getCell(1).getStringCellValue();
            String primeBusinessNameLower = currentRow.getCell(0).getStringCellValue().toLowerCase();
            String whoisBusinessNameLower = currentRow.getCell(1).getStringCellValue().toLowerCase();
            
            //#1 - LevenshteinDetailedDistance
            LevenshteinResults result = ldd.apply(primeBusinessNameLower, whoisBusinessNameLower);
            int distance = result.getDistance();
            
            Cell cellDistance = currentRow.createCell(2, CellType.NUMERIC);
            cellDistance.setCellValue(distance);
            
            //#2 - JaroWinklerDistance
            Double jwDistance = jwd.apply(primeBusinessNameLower, whoisBusinessNameLower);            
            Cell jwPercentage = currentRow.createCell(3, CellType.NUMERIC);            
            jwPercentage.setCellValue(Double.valueOf(df.format(jwDistance)));
            
            //#3 - Soundex
            int soundexDistance = soundex.difference(primeBusinessNameLower, whoisBusinessNameLower);            
            Cell cellSoundex = currentRow.createCell(4, CellType.NUMERIC);            
            cellSoundex.setCellValue(soundexDistance);
            
            /*//#4 - Refined Soundex
            int rsDistance = fs.difference(primeBusinessNameLower, whoisBusinessNameLower);            
            Cell cellFuzzy = currentRow.createCell(5, CellType.NUMERIC);            
            cellFuzzy.setCellValue(rsDistance);*/
            
            //#5 - Double Metaphone
            boolean matched = dm.isDoubleMetaphoneEqual(primeBusinessNameLower, whoisBusinessNameLower);            
            Cell cellDM = currentRow.createCell(6, CellType.BOOLEAN);            
            cellDM.setCellValue(matched);
            
            //#6 - N-gram
            //https://itssmee.wordpress.com/2010/06/28/java-example-of-n-gram/
            Double ngValue = ng.getSimilarity(primeBusinessNameLower, whoisBusinessNameLower, 2);
            Cell cellNgram = currentRow.createCell(7, CellType.NUMERIC);            
            cellNgram.setCellValue(Double.valueOf(df.format(ngValue)));
            
            
            /*System.out.println(" "
            		+ new Double(currentRow.getCell(2).getNumericCellValue()).intValue() + "\t"
            		+ currentRow.getCell(3).getNumericCellValue() + "\t"          		 
            		+ primeBusinessName + "\t\t" + whoisBusinessName);    */                                            

        }
        
        FileOutputStream outputStream = new FileOutputStream(DIST_FILE_NAME);
        workbook.write(outputStream);
        outputStream.close();
	}
}
