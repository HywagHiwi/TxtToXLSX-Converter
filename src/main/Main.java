package main;

import java.io.File;
import java.util.Scanner;

import javax.swing.JDialog;
import javax.swing.JFileChooser;
import javax.swing.JInternalFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;

import javafx.application.Application;
import javafx.application.Platform;
import javafx.stage.FileChooser;
import javafx.stage.FileChooser.ExtensionFilter;
import javafx.stage.Stage;
import javafx.stage.StageStyle;
import javafx.stage.Window;

/**
 * Hauptklasse der TxtToXLSX-Software.
 * @author snikelski
 *
 */
public class Main extends Application {

	@Override
	public void start(Stage primaryStage) throws Exception{
		primaryStage.initStyle(StageStyle.TRANSPARENT);
		
		FileChooser fileChooser = new FileChooser();
		fileChooser.setTitle("Auswahl der .txt-Datei.");
		try{
			fileChooser.setInitialDirectory(new File("K:/Auswertung Daten"));
		} catch(Exception e) {
			e.printStackTrace();
		}
		fileChooser.getExtensionFilters().addAll(new ExtensionFilter("Text-Dateien", "*.txt"));
		File selectedFile = fileChooser.showOpenDialog(primaryStage);
		if (selectedFile != null) {
			TxtReader txtReader = new TxtReader(selectedFile.getAbsolutePath());
			
			String filename = selectedFile.getName().substring(0, selectedFile.getName().length() -4);
			try{
				XlsxCreator xlsxCreator = new XlsxCreator(txtReader.getText(), filename);
				JOptionPane.showMessageDialog(new JInternalFrame(),
					    "Die Datei wurde erfolgreich generiert");
			} catch(WrongFormatException wfe) {
				JOptionPane.showMessageDialog(new JInternalFrame(), wfe.getMessage(), "FEHLER", JOptionPane.ERROR_MESSAGE);
				JOptionPane.showMessageDialog(new JInternalFrame(), "Die Datei wurde nicht generiert", "FEHLER", JOptionPane.ERROR_MESSAGE);
			}
			
			Platform.exit();
			System.exit(0);
			
		}
	
	}
	
	/**
	 * 
	 */
	public static void main(String[] args) {
		
		launch(args);
		
	}

}
