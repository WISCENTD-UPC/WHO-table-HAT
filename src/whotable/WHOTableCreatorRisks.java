package whotable;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.io.Writer;
import java.net.MalformedURLException;
import java.nio.charset.StandardCharsets;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FontUnderline;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

public class WHOTableCreatorRisks {
	
	
	
	public static void main (String[] args){
		
		String basepath = "HAT_webpage";
		
		char[] alphabet = "abcdefghijklmnopqrstuvwxyz".toUpperCase().toCharArray();
		
		String mainDescription = "\\$\\$MainDescription";
		String countryButtonsPattern = "\\$\\$?Countries";
		String countryModalPattern = "\\$\\$\\$\\?myModal";
		String countryButtonPattern = "\\$\\$\\$\\?myBtn";
		String countryCloseModalPattern = "\\$\\$\\$\\?closeModal";
		String countryModalTitlePattern = "\\$\\$\\$\\?title";
		String countryModalDescriptionPattern = "\\$\\$\\$\\?description";
		String countryModalProfilesPattern = "\\$\\$\\$\\?profiles";
		String countryModalDistributionsPattern = "\\$\\$\\$\\?distributions";
		String countryModalGuidelinesesPattern = "\\$\\$\\$\\?guidelineses";
		String countryModalSourcePicPattern = "\\$\\$\\$\\?sourcepic";
		String countryModalSourcePicAltPattern = "\\$\\$\\$\\?alt";
		
		String countryModalScript = "modal\\$\\$\\$";
		String countryBttnScript = "btn\\$\\$\\$";
		
		String countryDivsPattern = "\\$\\$\\$CountryDivs";
		String countryScriptsPattern = "\\$\\$\\$CountryScripts";
		
		HashMap<String, String> countryButtons = new HashMap<String,String>();
		
		FileInputStream file;
		FileInputStream file_modal;
		try {
			file = new FileInputStream(new File(basepath + "\\HAT_Webpage.xlsx" ));
		
			Workbook workbook = new XSSFWorkbook(file);

			Sheet sheet = workbook.getSheetAt(0);			 						
			
			String countryDivs = "";
			String countryScripts = "";
			
			BufferedReader br_page = new BufferedReader(new FileReader(new File(basepath + "\\resources\\table.html" )));
			String table = "";
			String l_page = "";
			while ((l_page = br_page.readLine()) != null){
				table += l_page + "\n";				
			}
			
			String main_description = "";
			FileInputStream fps = new FileInputStream(new File(basepath + "\\HAT_Webpage_Narrative_Risk.docx"));
			XWPFDocument docu = new XWPFDocument(fps);
			List<XWPFParagraph> data = docu.getParagraphs();
			for(XWPFParagraph p : data) {
				main_description += "\n<p";
//				System.err.println(p.getRuns().size());
				int f_size = 0;
				boolean font = false;
				for (XWPFRun r: p.getRuns()){
//					System.out.println(r.getFontFamily());					
					f_size = (r.getFontSize()!= -1)?r.getFontSize():12;					
					if (!font) main_description += " style=\"font-size:"+ f_size + "pt\">";
					font = true;
					if (r.isItalic())
						main_description += "<i>" + r.text() + "</i>";
					else if (r.isBold())
						main_description += "<b>" + r.text() + "</b>";
					else if (r.getUnderline() != UnderlinePatterns.NONE)
						main_description += "<u>" + r.text() + "</u>";
					else 
						main_description += r.text();
				}
				main_description += "</p>\n";				
//				System.out.println(p.getText());
//				System.out.println();
				
			}
			
//			System.out.println(main_description);
			
			
			for (int i=2; i<sheet.getLastRowNum(); i++){
				if (sheet.getRow(i) != null && (sheet.getRow(i)).getCell(0) != null){					

					BufferedReader br = new BufferedReader(new FileReader(new File(basepath + "\\resources\\modalCountryRisks.html" )));
					String divModal = "";
					String l = "";
					while ((l = br.readLine()) != null){
						divModal += l + "\n";				
					}
					//System.out.println(divModal);
					
					BufferedReader br_script = new BufferedReader(new FileReader(new File(basepath + "\\resources\\scriptModal.js" )));
					String script = "";
					String l_script = "";
					while ((l_script = br_script.readLine()) != null){
						script += l_script + "\n";				
					}
					
					
					String country_name = ((sheet.getRow(i)).getCell(0)!=null)?(sheet.getRow(i)).getCell(0).getStringCellValue():"";
					String country_code = ((sheet.getRow(i)).getCell(1)!=null)?(sheet.getRow(i)).getCell(1).getStringCellValue():"";
					//String country_description = ((sheet.getRow(i)).getCell(2)!=null)?(sheet.getRow(i)).getCell(2).getStringCellValue():"";
					
					//XSSFRichTextString rts =  (XSSFRichTextString)(sheet.getRow(i)).getCell(2).getRichStringCellValue();					
					//rts.getIndexOfFormattingRun(0);
					
					
					
//					CellStyle cs = ((sheet.getRow(i)).getCell(2)).getCellStyle();		
//					Font font = workbook.getFontAt(cs.getFontIndex());					
//					System.out.println(font.getItalic());
//					String fontName = font.getFontName();
//					System.out.println(fontName);
					String description = "";
//					if (rts.numFormattingRuns() == 0) {
//						description += rts.getString();
//					}
//					else{ 
//						for (int h = 0; h<rts.numFormattingRuns(); h++){
//							if (rts.getFontOfFormattingRun(h) != null && rts.getFontOfFormattingRun(h).getItalic()){
//								String txt = rts.getString();
//								description += "<i>" + txt.substring(rts.getIndexOfFormattingRun(h), rts.getIndexOfFormattingRun(h) + rts.getLengthOfFormattingRun(h)) +"</i>";
//							}
//							else if (rts.getFontOfFormattingRun(h) != null && rts.getFontOfFormattingRun(h).getBold()){
//								String txt = rts.getString();
//								description += "<b>" + txt.substring(rts.getIndexOfFormattingRun(h), rts.getIndexOfFormattingRun(h) + rts.getLengthOfFormattingRun(h)) +"</b>";
//							}
//							else if (rts.getFontOfFormattingRun(h) != null && (rts.getFontOfFormattingRun(h).getUnderline() != FontUnderline.NONE.getByteValue())){
//								String txt = rts.getString();
//								description += "<u>" + txt.substring(rts.getIndexOfFormattingRun(h), rts.getIndexOfFormattingRun(h) + rts.getLengthOfFormattingRun(h)) +"</u>";
//							}
//							else{
//								String txt = rts.getString();
//								description += txt.substring(rts.getIndexOfFormattingRun(h), rts.getIndexOfFormattingRun(h) + rts.getLengthOfFormattingRun(h));
//							}
//						}	
//					}
//					
//					description = description.replaceAll("\n", "<\\br>");
					
					
					//String country_profile = ((sheet.getRow(i)).getCell(3)!=null)?(sheet.getRow(i)).getCell(3).getStringCellValue():"";
					String distribution = ((sheet.getRow(i)).getCell(3)!=null)?(sheet.getRow(i)).getCell(3).getStringCellValue():"";
					//String national_guidelines = ((sheet.getRow(i)).getCell(5)!=null)?(sheet.getRow(i)).getCell(5).getStringCellValue():"";
					
					//System.out.println(country_name);
					String country_name_nospace = ((country_name.replaceAll(" ", "")).replaceAll("\\(", "")).replaceAll("\\)", "");
					country_name_nospace = ((country_name_nospace.replaceAll("ô", "o")).replaceAll("é", "e")).replaceAll("\\'", "");
					

					
					String countryFirstLetter  =  countryButtonsPattern.replaceAll("\\?", country_name.substring(0,1).toUpperCase());
					//System.out.println(countryFirstLetter);
					String country_button_html = "<button id=myBtn"+country_name_nospace+">"+country_name+"</button>";
					if (!countryButtons.containsKey(countryFirstLetter)){
						countryButtons.put(countryFirstLetter, country_button_html);
					}
					else{
						countryButtons.put(countryFirstLetter,countryButtons.get(countryFirstLetter)+country_button_html);
					}
					
					String myModalID = "myModal" + country_name_nospace;
					String myModalCloseID = "closeModal" + country_name_nospace;
					
					divModal = divModal.replaceAll(countryModalPattern, myModalID);
					divModal = divModal.replaceAll(countryCloseModalPattern, myModalCloseID);
					
					
//					//country profiles
//					String profile_html = "";
//					String[] profiles = country_profile.split("\n");
//					for (String p: profiles){
//						if (!p.contains("(")) continue;
//						String p1 = p.substring(0, p.indexOf("(")).trim();
//						String p2 = p.substring(p.indexOf("(")+1, p.indexOf(")"));						
//						//System.out.print(p1+" - "); System.out.println(p2);			
//						profile_html = profile_html + "<tr><td><a  target='_blank'  href='Country profiles/"+p2+"'>"+p1+"</a></td></tr>\n";
//						//System.out.println(profile_html);
//						
//					}
					
					//distribution
					String distribution_html = "";
					String[] distributions = distribution.split("\n");
					for (String p: distributions){
						if (!p.contains("(")){
							distribution_html = distribution_html + "<tr><td><i>"+p+"</i></td></tr>\n"; 
							continue;
						}
						String p1 = p.substring(0, p.indexOf("(")).trim();
						String p2 = p.substring(p.indexOf("(")+1, p.indexOf(")"));						
						//System.out.print(p1+" - "); System.out.println(p2);
						distribution_html = distribution_html + "<tr><td><a  target='_blank'  href='Risk/"+p2+"'>"+p1+"</a></td></tr>\n";
						//System.out.println(distribution_html);
						File tmp = new File (basepath+"\\Risk\\"+p2);
						if (!tmp.exists()) {
							System.err.println(p2+" file does not exist!");
						}
						else System.out.println(p2+" file exists!");
					}
					
					
					
//					//national_guidelines
//					String national_guidelines_html = "";
//					String[] national_guidelineses = national_guidelines.split("\n");
//					for (String p: national_guidelineses){
//						if (!p.contains("(")) continue;
//						String p1 = p.substring(0, p.indexOf("(")).trim();
//						String p2 = p.substring(p.indexOf("(")+1, p.indexOf(")"));						
//						//System.out.print(p1+" - "); System.out.println(p2);				
//						national_guidelines_html = national_guidelines_html + "<tr><td><a  target='_blank'  href='National guidelines/"+p2+"'>"+p1+"</a></td></tr>\n";
//						//System.out.println(national_guidelines_html);
//					}
					
					divModal = divModal.replaceAll(countryModalTitlePattern, country_name);
//					divModal = divModal.replaceAll(countryModalDescriptionPattern, description);
//					divModal = divModal.replaceAll(countryModalProfilesPattern, profile_html);
					divModal = divModal.replaceAll(countryModalDistributionsPattern, distribution_html);
//					divModal = divModal.replaceAll(countryModalGuidelinesesPattern, national_guidelines_html);
					countryDivs += divModal + "\n";
					
					script = script.replaceAll(countryButtonPattern, "myBtn"+country_name_nospace );
					script = script.replaceAll(countryModalPattern, myModalID);
					script = script.replaceAll(countryCloseModalPattern, myModalCloseID);
					script = script.replaceAll(countryModalScript, "modal"+country_name_nospace);
					script = script.replaceAll(countryBttnScript, "btn"+country_name_nospace);
					countryScripts += script + "\n";
					
					
					
					
					
					
				}
			}
			
			
			table = table.replaceAll(countryDivsPattern, countryDivs);
			table = table.replaceAll(countryScriptsPattern, countryScripts);
			
			for (char l:alphabet){
				String key = countryButtonsPattern.replaceAll("\\?", String.valueOf(l));
//				System.out.println(key);
				if (countryButtons.containsKey(key))
					table = table.replaceAll(key,countryButtons.get(key));
				else 
					table = table.replaceAll(key,"<br/>");
			}
			table = table.replaceAll(mainDescription, main_description);
			
//			System.out.println(table);
			
			Writer fstream = null;
			
			fstream = new OutputStreamWriter(new FileOutputStream(new File(basepath + "\\table.html")), StandardCharsets.UTF_8);
			BufferedWriter out = new BufferedWriter(fstream);
			out.write(table);
			
			out.close();
		  
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

}
