package helpers;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.AtomicMoveNotSupportedException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.util.function.IntConsumer;
import java.util.function.IntPredicate;

public final class SpreadsheetRowHelper {
    private SpreadsheetRowHelper() {
    }

    public static void processRows(
            Sheet sheet,
            Workbook workbook,
            File workbookFile,
            IntPredicate hasData,
            IntConsumer processRow) throws IOException {
        int rowIndex = 1;
        while (rowIndex <= sheet.getLastRowNum()) {
            Row row = sheet.getRow(rowIndex);
            if (row == null || !hasData.test(rowIndex)) {
                rowIndex++;
                continue;
            }

            processRow.accept(rowIndex);
            deleteRow(sheet, rowIndex);
            saveWorkbook(workbook, workbookFile.toPath());
        }
    }

    public static void ensureWorkbookClosed(File workbookFile) {
        File lockFile = new File(
                workbookFile.getParentFile(), "~$" + workbookFile.getName());
        if (lockFile.exists()) {
            throw new IllegalStateException(
                    "Close DCWorkbook.xlsm in Excel before running this program so completed rows can be saved. "
                            + "If Excel is closed and this lock file is stale, verify that first.");
        }
    }

    private static void deleteRow(Sheet sheet, int rowIndex) {
        int lastRowIndex = sheet.getLastRowNum();
        if (rowIndex < lastRowIndex) {
            sheet.shiftRows(rowIndex + 1, lastRowIndex, -1);
        } else {
            Row row = sheet.getRow(rowIndex);
            if (row != null) {
                sheet.removeRow(row);
            }
        }
    }

    private static void saveWorkbook(Workbook workbook, Path workbookPath) throws IOException {
        Path absolutePath = workbookPath.toAbsolutePath();
        Path temporaryPath = Files.createTempFile(
                absolutePath.getParent(), absolutePath.getFileName().toString(), ".tmp");
        try {
            try (OutputStream output = Files.newOutputStream(temporaryPath)) {
                workbook.write(output);
            }
            try {
                Files.move(temporaryPath, absolutePath,
                        StandardCopyOption.ATOMIC_MOVE, StandardCopyOption.REPLACE_EXISTING);
            } catch (AtomicMoveNotSupportedException e) {
                Files.move(temporaryPath, absolutePath, StandardCopyOption.REPLACE_EXISTING);
            }
        } finally {
            Files.deleteIfExists(temporaryPath);
        }
    }
}
