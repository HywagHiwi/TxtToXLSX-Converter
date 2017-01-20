package main;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class XlsxCreator {
	

	public XlsxCreator(List<String> text, String filename) throws WrongFormatException{
		Workbook wb = new org.apache.poi.xssf.usermodel.XSSFWorkbook();
		Sheet sheet = wb.createSheet("Tabelle1");
		
		CreationHelper createHelper = wb.getCreationHelper();
		
		Row header = sheet.createRow(0);
		header.createCell(0).setCellValue("Datum");
		header.createCell(1).setCellValue("Uhrzeit");
		header.createCell(2).setCellValue("Windrichtung");
		header.createCell(3).setCellValue("Windgeschwindigkeit");
		header.createCell(4).setCellValue("Temperatur");
		header.createCell(5).setCellValue("Niederschlag");
		header.createCell(6).setCellValue("relative Luftfeuchte");
		header.createCell(7).setCellValue("Luftdruck");
		header.createCell(8).setCellValue("Sättigungsdruck");
		header.createCell(9).setCellValue("Sättigungsdefizit");		
		header.createCell(10).setCellValue("Dampfdruck");
		header.createCell(11).setCellValue("Taupunkt");
		header.createCell(12).setCellValue("Windchill");
		header.createCell(13).setCellValue("Netto Radiometer");
		header.createCell(14).setCellValue("Globalstrahlung");
		header.createCell(15).setCellValue("Verdunstung");
		header.createCell(16).setCellValue("Pluvio");
		
		int i = 1; // Zaehlvariable fuer die Zeilenanzahl
		
		for(String s : text) { //Bedingung einfuegen (irgendwas mit Zeilen)
			
			if(!s.startsWith("Datum")) {
				
				String[] split = s.split(";");
				String[] date  = split[0].split(" "); //Datum und Uhrzeit separieren
				
				Row row = sheet.createRow(i); //Erschafft eine neue Reihe
				CellStyle cellStyle_Date = wb.createCellStyle();
				CellStyle cellStyle_Time = wb.createCellStyle();
				Cell cell;
				
				/*
				 * DATUM
				 */
				cellStyle_Date.setDataFormat(createHelper.createDataFormat().getFormat("dd.mm.yyyy"));
				cell = row.createCell(0);
				cell.setCellValue(date[0]);  
				cell.setCellStyle(cellStyle_Date);
				
				/*
				 * Uhrzeit
				 */
				cellStyle_Time.setDataFormat(createHelper.createDataFormat().getFormat("hh:mm:ss")); //Wird irgendwie noch nicht korrekt als Uhrzeit erkannt...
				cell = row.createCell(1);
				cell.setCellValue(date[1] + ":00");
				cell.setCellStyle(cellStyle_Time);
							
				/*
				 * Windrichtung
				 */
				row.createCell(2).setCellValue(changeDot(split[1]));
				
				/*
				 * Windgeschwindigkeit
				 */
				row.createCell(3).setCellValue(changeDot(split[2]));
				
				/*
				 * Temperatur
				 */
				row.createCell(4).setCellValue(changeDot(split[3]));
				
				/*
				 * Niederschlag
				 */
				row.createCell(5).setCellValue(changeDot(split[4]));
				
				/*
				 * relative Luftfeuchte
				 */
				row.createCell(6).setCellValue(changeDot(split[5]));
				
				/*
				 * Luftdruck
				 */
				row.createCell(7).setCellValue(changeDot(split[6]));
				
				/*
				 * Saettigungsdruck
				 */
				row.createCell(8).setCellValue(changeDot(split[7]));
				
				/*
				 * Saettigungsdefizit
				 */
				row.createCell(9).setCellValue(changeDot(split[8]));
				
				/*
				 * Dampfdruck
				 */
				row.createCell(10).setCellValue(changeDot(split[9]));
				
				/*
				 * Taupunkt
				 */
				row.createCell(11).setCellValue(changeDot(split[10]));
				
				/*
				 * Windchill
				 */
				row.createCell(12).setCellValue(changeDot(split[11]));
				
				/*
				 * Netto Radiometer
				 */
				row.createCell(13).setCellValue(changeDot(split[12]));
				
				/*
				 * Globalstrahlung
				 */
				row.createCell(14).setCellValue(changeDot(split[13]));
				
				/*
				 * Verdunstung
				 */
				row.createCell(15).setCellValue(changeDot(split[14]));
				
				/*
				 * Pluvio
				 */
				row.createCell(16).setCellValue(changeDot(split[15]));
				
				i++;
				
			} else{
				String[] split = s.split(";");
				if(split.length != 16) {
					throw new WrongFormatException("Falsche Anzahl von Parametern!");
				} else {
					if(!split[0].equals("Datum")){
						throw new WrongFormatException("Das Datum ist nicht an erster Stelle");
					}
					if(!split[1].equals("Windrichtung")){
						throw new WrongFormatException("Die Windrichtung ist nicht an zweiter Stelle");
					}
					if(!split[2].equals("Windgeschwindigkeit")){
						throw new WrongFormatException("Die Windgeschwindigkeit ist nicht an passender Position");
					}
					if(!split[3].equals("Temperatur")){
						throw new WrongFormatException("Die Temperatur ist nicht an passender Position");
					}
					if(!split[4].equals("Niederschlag")){
						throw new WrongFormatException("Der Niederschlag ist nicht an passender Position");
					}
					if(!split[5].equals("relative Luftfeuchte")){
						throw new WrongFormatException("Die relative Luftfeuchte ist nicht an passender Position");
					}
					if(!split[6].equals("Luftdruck")){
						throw new WrongFormatException("Der Luftdruck ist nicht an passender Position");
					}
					if(!split[7].equals("Sättigungsdruck")){
						throw new WrongFormatException("Der Sättigungsdruck ist nicht an passender Position");
					}
					if(!split[8].equals("Sättigungsdefizit")){
						throw new WrongFormatException("Das Sättigungsdefizit ist nicht an passender Position");
					}
					if(!split[9].equals("Dampfdruck")){
						throw new WrongFormatException("Der Dampfdruck ist nicht an passender Position");
					}
					if(!split[10].equals("Taupunkt")){
						throw new WrongFormatException("Der Taupunkt ist nicht an passender Position");
					}
					if(!split[11].equals("Windchill")){
						throw new WrongFormatException("Der Windchill ist nicht an passender Position");
					}
					if(!split[12].equals("Netto Radiometer")){
						throw new WrongFormatException("Der Netto Radiometer ist nicht an passender Position");
					}
					if(!split[13].equals("Globalstrahlung")){
						throw new WrongFormatException("Die Globalstrahlung ist nicht an passender Position");
					}
					if(!split[14].equals("Verdunstung")){
						throw new WrongFormatException("Die Verdunstung ist nicht an passender Position");
					}
					if(!split[15].equals("Pluvio")){
						throw new WrongFormatException("Der Pluvio ist nicht an passender Position");
					}
				}		
						
			}
		}
		
		
		
		FileOutputStream fileOut;
		try {
			fileOut = new FileOutputStream(filename + ".xlsx");
			wb.write(fileOut);
			fileOut.close();
			wb.close();
		} catch (FileNotFoundException e) {
			System.out.println("Die Datei konnte nicht erzeugt werden.");
		} catch (IOException e) {
			System.out.println("Die Datei konnte nicht beschrieben werden.");
		}		
	}

	/**
	 * Methode zur Umwandlung der Punkte in Kommata (fuer Access). NICHT aufs Datum anwenden.
	 * @param s
	 * @return
	 */
	private String changeDot(String s) {
		String es = s.replace(".", ",");
		if(es.equals("")) {
			es = "*";
		}
		return es;
	}
	
}
