import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;
import org.json.simple.parser.ParseException;

import java.io.*;

import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.time.LocalDate;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.HashMap;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

import static org.apache.log4j.spi.Configurator.NULL;

public class act_list_compare {
    static String directoryPath = "./";
    static String outputDirectoryPath = "./output_date/";
    static String inputDirectoryPath = "./input_date/";
    static String archiveDirectoryPath = "./out/archive/";
    static String outXlsDirectory = "./out/xls/";
    public static void archiver() {
        File outputDirectory = new File(outputDirectoryPath);
        File inputDirectory = new File(inputDirectoryPath);
        File zipFile = new File(dirName());
        File[] outputDirectoryFiles = outputDirectory.listFiles();
        File[] inputDirectoryFiles = inputDirectory.listFiles();

        if (inputDirectoryFiles != null) {
            for (File file : inputDirectoryFiles) {
                String fileName = file.getName();

                String tempOurDir = inputDirectoryPath + fileName;
                String tempArchDir = archiveDirectoryPath + fileName;

                Path sourcePath = Path.of(tempOurDir);
                Path archivePath = Path.of(tempArchDir);

                try {
                    Files.copy(sourcePath, archivePath, StandardCopyOption.REPLACE_EXISTING);
                } catch (Exception e) {
                    throw new RuntimeException(e);
                }
            }
        }

        if (outputDirectoryFiles != null) {
            for (File file : outputDirectoryFiles) {
                String fileName = file.getName();
                String tempOurDir = outputDirectoryPath + fileName;
                String tempArchDir = archiveDirectoryPath + fileName;
                String tempZip = archiveDirectoryPath + zipFile;
                String directoryPath = "./держифайлик.xlsx";

                Path sourcePath = Path.of(tempOurDir);
                Path archivePath = Path.of(tempArchDir);
                Path zipPatch = Path.of(tempZip);
                Path directoryzipPath = Path.of(directoryPath);
                try {
                    Files.copy(sourcePath, archivePath, StandardCopyOption.REPLACE_EXISTING);
                    Files.copy(sourcePath, directoryzipPath, StandardCopyOption.REPLACE_EXISTING);
                    System.out.println("Copy to "+ directoryzipPath + " " + fileName);
                    System.out.println("Successful");
                    Files.copy(Path.of(String.valueOf(zipFile)), zipPatch, StandardCopyOption.REPLACE_EXISTING);
                    System.out.println("Copy to "+ archiveDirectoryPath + " " + zipFile);
                    System.out.println("Successful");
                } catch (IOException e) {
                    throw new RuntimeException(e);
                }
            }
        }
    }
    public static void cleaner() {
        File outXlsFolder = new File(outXlsDirectory);
        File inputDirectoryFolder = new File(inputDirectoryPath);
        File outputDirectoryFolder = new File(outputDirectoryPath);
        File zipFile = new File(dirName());

        File[] filesXlsCleaner = outXlsFolder.listFiles();
        File[] inputDirectoryCleaner = inputDirectoryFolder.listFiles();
        File[] outputDirectoryCleaner = outputDirectoryFolder.listFiles();

        if (filesXlsCleaner != null) {
            for (File file : filesXlsCleaner) {
                file.delete();
            }
            System.out.println("./out/xls directory is clean");
        }

        if (inputDirectoryCleaner != null) {
            for (File file : inputDirectoryCleaner) {
                file.delete();
            }
            System.out.println("./input_date directory is clean");
        }

        if (outputDirectoryCleaner != null) {
            for (File file : outputDirectoryCleaner) {
                file.delete();
            }
            System.out.println("./output_date directory is clean");
        }

        if (zipFile.exists()) {
            if (zipFile.delete()) {
                System.out.println(".zip is deleted");
            }
        }
    }
    public static void copyFilesFromFolder(File folder) throws IOException, InterruptedException {
        File[] folderEntries = folder.listFiles();

        for (File entry : folderEntries) {
            if (entry.isDirectory()) {
                copyFilesFromFolder(entry);
                continue;
            }

            String fileName = entry.getName();
            System.out.println("inpute file: " + fileName);

            DateTimeFormatter formatter = DateTimeFormatter.ISO_LOCAL_TIME;
            LocalTime time = LocalTime.now();

            //добавление времени в имя файла
//            String delNano = formatter.format(time).replaceAll("\\.[^.]*", "");
//            String toFormat = delNano.replaceAll(":", "-");

            File source = new File("./input_date/" + fileName);
            //ждать 1 сек, включать если больше одного файла в инпуте и добалено время в имя файла
//            try {
//                TimeUnit.SECONDS.sleep(1);
//            } catch (InterruptedException e) {
//                throw new RuntimeException(e);
//            }
            File dest = new File("./output_date/" + LocalDate.now() + "T" + ".xlsx");

            try {
                Files.copy(source.toPath(), dest.toPath(), StandardCopyOption.REPLACE_EXISTING);

            } catch (IOException e) {
                throw new RuntimeException(e);
            }

            System.out.println("output file: " + dest);
        }
    }
    public static int processFilesFromFolderClients(File folder) throws IOException, InterruptedException {
        File[] folderEntries = folder.listFiles();

        for (File entry : folderEntries) {
            if (entry.isDirectory()) {
                copyFilesFromFolder(entry);
                continue;
            }

            DataFormatter formatter = new DataFormatter();
            String fileName = entry.getName();
            System.out.println(fileName);

            // Read XSL file
            FileInputStream inputStreamClient = new FileInputStream(new File("./out/xls/"+fileName));
            FileInputStream inputStreamAct = new FileInputStream(new File("./output_date/" + LocalDate.now() + "T" + ".xlsx"));

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

            FileOutputStream outputStream = new FileOutputStream(new File("./output_date/" + LocalDate.now() + "T" + ".xlsx"));
            workbookAct.write(outputStream);
            workbookAct.close();

            outputStream.close();
            inputStreamAct.close();
            inputStreamClient.close();
        }
        return 0;
    }
    public static void MaxValueByDuplicate() {
        // Пример данных для таблицы (двумерного массива)
        int[][] table = {
                {1, 5},
                {1, 10},
                {1, 10},
                {2, 15},
                {2, 20},
                {2, 5},
                {2, 25},
                {3, 15},
                {3, 25},
                {3, 30}
        };

        // Создаем отображение для хранения наибольших значений из второго столбца
        Map<Integer, Integer> maxValueMap = new HashMap<>();

        // Проходим по каждой строке таблицы
        for (int i = 0; i < table.length; i++) {
            // Значение в первом столбце
            int firstColumnValue = table[i][0];
            // Значение во втором столбце
            int secondColumnValue = table[i][1];

            // Проверяем, есть ли уже значение в отображении для данного значения первого столбца
            if (maxValueMap.containsKey(firstColumnValue)) {
                // Если есть, то сравниваем текущее значение второго столбца с уже имеющимся наибольшим значением
                if (secondColumnValue > maxValueMap.get(firstColumnValue)) {
                    // Если текущее значение больше, обновляем наибольшее значение в отображении
                    maxValueMap.put(firstColumnValue, secondColumnValue);
                }
            } else {
                // Если нет, добавляем значение в отображение
                maxValueMap.put(firstColumnValue, secondColumnValue);
            }
        }

        // Выводим результат
        System.out.println("Наибольшие значения из второго столбца для повторяющихся значений первого столбца:");
        for (Map.Entry<Integer, Integer> entry : maxValueMap.entrySet()) {
            System.out.println(entry.getKey() + " -> " + entry.getValue());
        }
    }
    public static String dirName() {
        String searchExtension = ".zip";
        String filePath = NULL;
        File directory = new File(directoryPath);
        File[] files = directory.listFiles();

        if (files != null) {
            for (File file : files) {
                if (file.isFile() && file.getName().endsWith(searchExtension)) {
                    filePath = file.getPath();
                }
            }
        }

        return filePath;
    }
    private static final Pattern NUMBER_PATTERN = Pattern.compile("[-]?[0-9]+(.[0-9]+)?");
    public static void slider(String t, int indx, XSSFSheet sheet) {
        XSSFRow row = sheet.createRow(indx);
        XSSFCell cell = row.createCell(0, CellType.STRING);
        XSSFCell cell2 = row.createCell(1, CellType.STRING);

        String replacer = t.replaceAll(" \n\n","; ");
        String replacer0 = replacer.replaceAll("  "," ");
        String replacer2 = replacer0.replaceAll("\n\n","; ");
        String replacer333 = replacer2.replaceAll(" \n","; ");
        String replacer3 = replacer333.replaceAll("\n","; ");
        String replacer44 = replacer3.replaceAll("Номер заказа СДЭК ", "");
        String replacer4 = replacer44.replaceAll("Накладная успешно создана с номером: ", "");
        String replacer55 = replacer4.replaceAll(" ", ";");
        String replacer6 = replacer55.replaceAll(";;", "; ");
        String replacer7 = replacer6.replaceAll(";", "; ");
        String replacer8 = replacer7.replaceAll("  ", " ");
        String replacer9 = replacer8.replaceAll("; ;", "; ");
        String Nakladnaya = replacer55.substring(0, 10);
        System.out.println(replacer9);
        cell.setCellValue(Nakladnaya);
        if (replacer9.length() > 10) {
            String[] split = replacer9.split(";");
            if (split.length > 1) {
                String price = split[1].split(" ")[1];
                Matcher matcher = NUMBER_PATTERN.matcher(price);
                while (matcher.find()) {
                    cell2.setCellValue(matcher.group());
                }
            }
        }
    }
    public static void processFilesFromFolder(File zipFile) throws IOException, ParseException {
        ZipInputStream zipInputStream = new ZipInputStream(new FileInputStream(zipFile));
        ZipEntry entry;

        InputStreamReader inputStreamReader = new InputStreamReader(zipInputStream);
        BufferedReader bufferedReader = new BufferedReader(inputStreamReader);

        while ((entry = zipInputStream.getNextEntry()) != null) {
            String fileName = entry.getName();
            System.out.println(fileName);

            JSONParser parser = new JSONParser();
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = workbook.createSheet();
            int indx = 0;

            JSONObject jsonObject = (JSONObject) parser.parse(bufferedReader);
            JSONArray msg = (JSONArray) jsonObject.get("data");

            if (!entry.isDirectory() && entry.getName().endsWith(".json")) {
                for (Object o : msg) {
                    JSONObject json_obj = (JSONObject) o;
                    JSONArray m = (JSONArray) json_obj.get("m");
                    String t1 = (String) json_obj.get("t");
                    if (t1.contains("Номер заказа СДЭК") || t1.contains("Накладная успешно")) {
                        slider(t1, indx, sheet);
                        indx++;
                    }
                    for (Object mObj : m) {
                        JSONObject mObjJson = (JSONObject) mObj;
                        String t = (String) mObjJson.get("t");
                        if (t.contains("Номер заказа СДЭК") || t.contains("Накладная успешно")) {
                            slider(t, indx, sheet);
                            indx++;
                        }
                    }
                }
                FileOutputStream outputStream = new FileOutputStream("out/xls/" + fileName + ".xlsx");
                workbook.write(outputStream);
            }
        }
        bufferedReader.close();
        zipInputStream.close();
    }
    public static void main(String[] args) throws IOException, ParseException, InterruptedException {
        File zipFile = new File(dirName()); //path указывает на директорию
        File dir = new File("./input_date/");
        processFilesFromFolder(zipFile);

        //MaxValueByDuplicate(); //тут бы обработать таблицу на повторения, но сценарий супередкий

        copyFilesFromFolder(dir);
        System.out.println("act copy complete");
        File dirClients = new File("./out/xls/");
        processFilesFromFolderClients(dirClients);

        archiver();
        cleaner();

        System.out.println("Program completed successfully");
    }
}