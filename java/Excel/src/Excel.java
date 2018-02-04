import java.io.*;
import java.util.*;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.*;

public class Excel {
    private HSSFWorkbook workBook;
    private HSSFSheet sheet;
    private HSSFCellStyle style;
    private static final short FONT_SIZE = 11;

    public void createFile(String[] header, String[][] data) throws IOException {
        workBook = new HSSFWorkbook();
        sheet = workBook.createSheet("シート名"); // シート名を指定する
        style = createCellStyle();
        writeHeader(header);
        write(data);
        OutputStream out = new FileOutputStream("ファイル名"); // ファイル名を指定する
        workBook.write(out);
        out.close();
    }

    private void writeHeader(String[] header) throws IOException {
        HSSFRow row = sheet.createRow(0);
        for (short i = 0; i < header.length; i++) {
            HSSFCell cell = row.createCell(i);
            cell.setEncoding(HSSFCell.ENCODING_UTF_16);
            cell.setCellValue(header[i]);
            cell.setCellStyle(style);
        }
        sheet.createFreezePane(0, 1, 0, 1);
    }

    private void write(String[][] data) throws IOException {
        short rowCount = 1;
        for (short i = 0; i < data.length; i++) {
            HSSFRow row = sheet.createRow(rowCount++);
            for (short j = 0; j < data.length; j++) {
                HSSFCell cell = row.createCell(j);
                cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                cell.setCellValue(data[i][j]);
                cell.setCellStyle(style);
            }
        }
    }

    private HSSFCellStyle createCellStyle() {
        HSSFFont font = workBook.createFont();
        font.setFontHeightInPoints(FONT_SIZE);
        font.setFontName("ＭＳ Ｐゴシック");
        HSSFCellStyle style = workBook.createCellStyle();
        style.setFont(font);
        style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        style.setBottomBorderColor(HSSFColor.BLACK.index);
        style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        style.setLeftBorderColor(HSSFColor.BLACK.index);
        style.setBorderRight(HSSFCellStyle.BORDER_THIN);
        style.setRightBorderColor(HSSFColor.BLACK.index);
        style.setBorderTop(HSSFCellStyle.BORDER_THIN);
        style.setTopBorderColor(HSSFColor.BLACK.index);
        return style;
    }

    public static void main(String[] args) {
        String[] header = new String[]{"年月日", "データ"};
        List list = new ArrayList();
        list.add(new String[]{"20040101", "500円"});
        list.add(new String[]{"20040102", "1000円"});
        String[][] data = new String[list.size()][];
        try {
            new Excel().createFile(header, (String[][])list.toArray(data));
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
