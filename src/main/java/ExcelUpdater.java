import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Scanner;

public class ExcelUpdater {

    public static void main(String[] args) {
        Scanner scanner = new Scanner(System.in);

        // Excelファイルのパス入力
        String filePath = "C:\\Users\\G-2020001\\Downloads\\出力データー在庫管理システム.xlsm";

        // 列タイトルの配列
        String[] columns = {"No.", "品番", "品名", "サイズ", "色", "在庫場所", "原価", "売価", "仕入れ先", "JANコード", "昨日在庫", "現在在庫"};

        // 入力するデータの内容を入力
        System.out.println("以下の内容を入力してください：");
        
        System.out.println("品番：");
        String 品番 = scanner.next();
        
        System.out.println("品名：");
        String 品名 = scanner.next();
        
        System.out.println("サイズ：");
        String サイズ = scanner.next();
        
        System.out.println("色：");
        String 色 = scanner.next();
        
        System.out.println("在庫場所：");
        String 在庫場所 = scanner.next();
        
        System.out.println("原価：");
        double 原価 = scanner.nextDouble();
        
        System.out.println("売価：");
        double 売価 = scanner.nextDouble();
        
        System.out.println("仕入れ先：");
        String 仕入れ先 = scanner.next();
        
        System.out.println("JANコード：");
        String JANコード = scanner.next();
        
        System.out.println("昨日在庫：");
        int 昨日在庫 = scanner.nextInt();
        
        System.out.println("現在在庫：");
        int 現在在庫 = scanner.nextInt();

        try {
            // 指定されたExcelファイルを読み込む
            FileInputStream fis = new FileInputStream(new File(filePath));
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheetAt(0); // 最初のシートを読み取る

            // 列タイトル行を取得（仮に1行目とする）
            Row headerRow = sheet.getRow(0);

            // No.列のインデックスを探す
            int noColumnIndex = -1;
            for (int i = 0; i < headerRow.getPhysicalNumberOfCells(); i++) {
                if (headerRow.getCell(i).getStringCellValue().equals("No.")) {
                    noColumnIndex = i;
                    break;
                }
            }

            if (noColumnIndex == -1) {
                System.out.println("'No.' 列が見つかりません！");
                return;
            }

            // No.列の最大値を探し、実際にデータが入っている最後の行を探す
            int maxNo = 0;
            int lastRowIndex = 0;  // 2行目から開始（データ行）

            // テーブルにタイトル行だけが含まれているかチェック、もしそうなら lastRowIndex を 1 に設定
            if (sheet.getPhysicalNumberOfRows() > 1) {
                // テーブルにデータ行が含まれている場合
                for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
                    Row row = sheet.getRow(i);
                    if (row != null) {
                        Cell cell = row.getCell(noColumnIndex);
                        if (cell != null && cell.getCellType() == CellType.NUMERIC) {
                            int currentNo = (int) cell.getNumericCellValue();
                            if (currentNo > maxNo) {
                                maxNo = currentNo;
                                lastRowIndex = i;  // 最後の有効な行のインデックスを更新
                            }
                        }
                    }
                }
            }

            // 最後の有効な行の次の行にデータを挿入
            Row newRow = sheet.createRow(lastRowIndex + 1);

            // No.列の値を自動的に増加
            newRow.createCell(noColumnIndex).setCellValue(maxNo + 1);  // 'No.' は現在の最大値 + 1

            // 他の列にデータを入力
            newRow.createCell(1).setCellValue(品番);
            newRow.createCell(2).setCellValue(品名);
            newRow.createCell(3).setCellValue(サイズ);
            newRow.createCell(4).setCellValue(色);
            newRow.createCell(5).setCellValue(在庫場所);
            newRow.createCell(6).setCellValue(原価);
            newRow.createCell(7).setCellValue(売価);
            newRow.createCell(8).setCellValue(仕入れ先);
            newRow.createCell(9).setCellValue(JANコード);
            newRow.createCell(10).setCellValue(昨日在庫);
            newRow.createCell(11).setCellValue(現在在庫);

            System.out.println("データがExcelファイルに追加されました！");

            // 更新された内容をファイルに書き込む
            fis.close();
            FileOutputStream fileOut = new FileOutputStream(new File(filePath));
            workbook.write(fileOut);
            fileOut.close();
            workbook.close();

        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            scanner.close();
        }
    }
}
