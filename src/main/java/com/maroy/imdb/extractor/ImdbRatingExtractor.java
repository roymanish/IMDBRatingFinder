package com.maroy.imdb.extractor;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FilenameFilter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.codehaus.jettison.json.JSONException;
import org.codehaus.jettison.json.JSONObject;
import org.springframework.web.client.RestTemplate;

public class ImdbRatingExtractor {

	private static String IMDB_REST_URL = "http://www.omdbapi.com/?t=<<KEY>>&y=&plot=short&r=json";
	
	public static void main(String[] args) {

		String url = "";
		String path = "C:\\Personal\\Personal\\Movies";
		ImdbRatingExtractor extractor = new ImdbRatingExtractor();
		String[] movies = extractor.getListOfMovieNames(path);
		
		List<JSONObject> ratingsList = new ArrayList<JSONObject>();
		for(String key : movies){
			
			url = IMDB_REST_URL;
			url = url.replace("<<KEY>>", key);
			
			RestTemplate restTemplate = new RestTemplate();
			String response = restTemplate.getForObject(url, String.class);
			JSONObject obj = null;
			try {
				obj = new JSONObject(response);
				
				if(obj.get("Response") != null && obj.get("Response").equals("False"))
					System.out.println("Invalid Movie" + key);
				else{
					ratingsList.add(obj);
				}
			} catch (JSONException e) {
				e.printStackTrace();
			}
		}
		
		try {
			extractor.writeResultsToFile(ratingsList);
		} catch (IOException e) {
			e.printStackTrace();
		} catch (JSONException e) {
			e.printStackTrace();
		}
	}
	
	private String[] getListOfMovieNames(String path){
		
		File file = new File(path);
		String[] directories = file.list(new FilenameFilter() {
			
			@Override
			public boolean accept(File dir, String name) {
				return true;
			}
		});
		return directories;
	}

	private void writeResultsToFile(List<JSONObject> response) throws IOException,JSONException{
		
		FileInputStream file = new FileInputStream(new File("RatingsTemplate.xls"));
        
		//Get the workbook instance for XLS file 
		HSSFWorkbook workbook = new HSSFWorkbook(file);
		 
		//Get first sheet from the workbook
		HSSFSheet sheet = workbook.getSheetAt(0);
		
		int i=1;
		for(JSONObject json : response){
			
			Row row = sheet.createRow(i++);
			
			Cell nameCell = row.createCell(0);
			nameCell.setCellValue(json.getString("Title"));
			
			Cell ratingCell = row.createCell(1);
			ratingCell.setCellValue(json.getString("imdbRating"));
		}
		
		file.close();
	    FileOutputStream out = 
	        new FileOutputStream(new File("RatingsTemplate.xls"));
	    workbook.write(out);
	    out.close();
	}
}
