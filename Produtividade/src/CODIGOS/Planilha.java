/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package CODIGOS;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JOptionPane;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author irmoura
 */
public class Planilha {
    
//    public static String[] PS = new String[36];
    public static String[] PS;
    public static int count, count2, num_linhas, num_colunas;
    
    public static void lerPlanilha(String arquivo){
        
        FileInputStream fisPlanilha = null;
        
        try {
            
            //"\\\\HP-G42\\Users\\Public\\GA_XLSX" - REDE
            //"C:\\GA_XLSX" - LOCAL
            
            File dir = new File("C:\\PRODUTIVIDADE");
            File file = new File(dir,arquivo+".xlsx");
            fisPlanilha = new FileInputStream(file);
            
            /*CRIA UM WORKBOOK = PLANILHA TODA COM TODAS AS ABAS*/
            XSSFWorkbook workbook = new XSSFWorkbook(fisPlanilha);
            
            /*RECUPERAMOS APENAS A PRIMEIRA ABA OU PRIMEIRA PLANILHA*/
            XSSFSheet sheet = workbook.getSheetAt(0);
            
            /*RETORNA TODAS AS LINHAS DA PLANILHA 0 */
            Iterator<Row> rowIterator = sheet.iterator();
            
            /*VARRE TODAS AS LINHAS DA PLANILHA 0*/
            while(rowIterator.hasNext()){
                
                //recebe cada linha da planilha
                Row row = rowIterator.next();
                
                //pegamos todas as celulas desta linha
                Iterator<Cell> cellIterator = row.iterator();
                
                //varremos todas as celulas da linha atual
                while(cellIterator.hasNext()){
                    
                    /*OBTEM AUTOMATICAMENTE O NUMERO TOTAL DE CELULAS*/
                    count2++;
                    if(count2 == 1){
                        num_colunas = row.getPhysicalNumberOfCells(); // Número de colunas
                        num_linhas = sheet.getPhysicalNumberOfRows(); // Número de linhas
                        PS = new String[num_linhas*num_colunas];
                        System.out.println(PS.length+" células encontradas.");
                    }
                    /*OBTEM AUTOMATICAMENTE O NUMERO TOTAL DE CELULAS*/
                    
                    //criamos uma celula
                    Cell cell = cellIterator.next();
                    
                    switch(cell.getCellType()){
                        
                        case Cell.CELL_TYPE_BLANK:
                            PS[count] = ""+cell.getStringCellValue();
                            count++;
                            break;
                        case Cell.CELL_TYPE_BOOLEAN:
                            PS[count] = ""+cell.getBooleanCellValue();
                            count++;
                            break;
                        case Cell.CELL_TYPE_ERROR:
                            PS[count] = ""+cell.getErrorCellValue();
                            count++;
                            break;
                        case Cell.CELL_TYPE_FORMULA:
                            PS[count] = ""+cell.getCellFormula();
                            count++;
                            break;
                        case Cell.CELL_TYPE_STRING:
                            PS[count] = ""+cell.getStringCellValue();
                            count++;
                            break;
                        case Cell.CELL_TYPE_NUMERIC:
                            PS[count] = ""+cell.getNumericCellValue();
                            PS[count] = PS[count].replace(".0", "");
                            count++;
                            break;
                        
                    }
                    
                }
                
            }
            
        } catch (FileNotFoundException ex) {
              JOptionPane.showMessageDialog(null,"Arquivo "+arquivo+" não encontrado.","Warning",JOptionPane.WARNING_MESSAGE);
              System.exit(0);
        } catch (IOException ex) {
            Logger.getLogger(Planilha.class.getName()).log(Level.SEVERE, null, ex);
        } finally {
            try {
                fisPlanilha.close();
            } catch (IOException ex) {
                Logger.getLogger(Planilha.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
        
    }
    
}