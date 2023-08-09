import java.io.IOException;
import java.io.File;
import java.nio.file.Files;
import java.time.LocalDate;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.concurrent.TimeUnit;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;

/*
Программа читает каталог dir
Копирует файл из каталога dir в каталог dest, меняя имя файла на текущую дату и время
Сравнивает две таблицы по первому столбцу, сравнивая числа во втором столбце.
Если разница попадает в условия, то заполняются последующие столбики
 */

public class Main {
    public static void processFilesFromFolder(File folder) throws IOException, InterruptedException {
        File[] folderEntries = folder.listFiles();

        for (File entry : folderEntries) {
            if (entry.isDirectory()) {
                processFilesFromFolder(entry);
                continue;
            }

            String fileName = entry.getName();
            System.out.println(fileName);

            DateTimeFormatter formatter = DateTimeFormatter.ISO_LOCAL_TIME;
            LocalTime time = LocalTime.now();

            //добавление времени в имя файла
//            String delNano = formatter.format(time).replaceAll("\\.[^.]*", "");
//            String toFormat = delNano.replaceAll(":", "-");

            File source = new File("C:\\Users\\Ilya\\IdeaProjects\\act_list_compare\\input_date\\" + fileName);
            //ждать 1 сек, включать если больше одного файла в инпуте и добалено время в имя файла
//            try {
//                TimeUnit.SECONDS.sleep(1);
//            } catch (InterruptedException e) {
//                throw new RuntimeException(e);
//            }
            File dest = new File("C:\\Users\\Ilya\\IdeaProjects\\act_list_compare\\output_date\\" + LocalDate.now() + "T" + ".xlsx");
            Files.copy(source.toPath(), dest.toPath());

            //удаление файла исходника
//            if (source.delete()){
//                System.out.println("source файл был удален");
//            }
//            else System.out.println("Файл source не был найден");
        }
    }

    public static int processFilesFromFolderClients(File folder) throws IOException, InterruptedException {
        File[] folderEntries = folder.listFiles();

        for (File entry : folderEntries) {
            if (entry.isDirectory()) {
                processFilesFromFolder(entry);
                continue;
            }

            DataFormatter formatter = new DataFormatter();
            String fileName = entry.getName();
            System.out.println(fileName);

            // Read XSL file
            FileInputStream inputStreamClient = new FileInputStream(new File("C:\\Users\\Ilya\\IdeaProjects\\act_list_compare\\xls\\"+fileName));
            FileInputStream inputStreamAct = new FileInputStream(new File("C:\\Users\\Ilya\\IdeaProjects\\act_list_compare\\output_date\\" + LocalDate.now() + "T" + ".xlsx"));

            // Get the workbook instance for XLS file
            XSSFWorkbook workbookClient = new XSSFWorkbook(inputStreamClient);
            XSSFWorkbook workbookAct = new XSSFWorkbook(inputStreamAct);

            // Get first sheet from the workbook
            XSSFSheet sheetClient = workbookClient.getSheetAt(0);
            XSSFSheet sheetAct = workbookAct.getSheetAt(0);

            int p = 0;

            for (int i = 0; i < sheetAct.getLastRowNum() + 1; i++) {

                XSSFCell cellX = sheetAct.getRow(i).getCell(0);
                XSSFCell cellX1 = sheetAct.getRow(i).getCell(1);

                String formatX = formatter.formatCellValue(cellX);
                String formatX1 = formatter.formatCellValue(cellX1);

                int nakl1 = Integer.parseInt(String.valueOf(formatX));
                int sum1;

                if (formatX1.length() == 0) {
                    sum1 = -9999999;
                } else {
                    sum1 = Integer.parseInt(String.valueOf(formatX1));
                }

                for (int y = 0; y < sheetClient.getLastRowNum() + 1; y++) {

                    XSSFCell cellY = sheetClient.getRow(y).getCell(0);
                    XSSFCell cellY1 = sheetClient.getRow(y).getCell(1);

                    String formatY = formatter.formatCellValue(cellY);
                    String formatY1 = formatter.formatCellValue(cellY1);

                    int nakl2;

                    try {
                        nakl2 = Integer.parseInt(String.valueOf(formatY));
                    } catch (NumberFormatException e) {
                       break;
                    }

                    int sum2;

                    if (formatY1.length() == 0) {
                        sum2 = -9999999;
                    } else {
                        try {
                            sum2 = Integer.parseInt(String.valueOf(formatY1));
                        } catch (NumberFormatException e) {
                            break;
                        }
                    }

                    if (nakl1 - nakl2 == 0) {
                        XSSFRow rowAct = sheetAct.getRow(i);
                        XSSFCell cellAct = rowAct.createCell(3, CellType.STRING);
                        XSSFCell cellActSum = rowAct.createCell(2, CellType.STRING);
                        int raznisa = (sum2 - sum1);

                        if (raznisa < 50) {
                            p++;
                            raznisa = Math.abs(raznisa) + 70;
                            System.out.print(nakl1 + " изменилась: ");
                            System.out.println(raznisa);

                            cellActSum.setCellValue(sum2);
                            cellAct.setCellValue(raznisa);
                        } else {
                            System.out.println(nakl1 + " не изменилась");

                            cellActSum.setCellValue(sum2);
                            cellAct.setCellValue(0);
                        }
                    }
                }
            }

            FileOutputStream outputStream = new FileOutputStream(new File("C:\\Users\\Ilya\\IdeaProjects\\act_list_compare\\output_date\\" + LocalDate.now() + "T" + ".xlsx"));
            workbookAct.write(outputStream);
            workbookAct.close();
        }
        return 0;
    }

    public static void main(String[] args) throws IOException, InterruptedException {
        File dir = new File("C:\\Users\\Ilya\\IdeaProjects\\act_list_compare\\input_date\\");
        processFilesFromFolder(dir); //копирования файлов из директорий
        System.out.println("act copy complete");

        File dirClients = new File("C:\\Users\\Ilya\\IdeaProjects\\act_list_compare\\xls\\");
        processFilesFromFolderClients(dirClients);

    }
}
