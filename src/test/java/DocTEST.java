import com.utils.DocPOI.ExcelReader;
import com.utils.DocPOI.ExcelWriter;
import com.utils.DocPOI.WordWriter;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblBorders;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.List;

import static com.utils.Main.PATH;

public class DocTEST {

    @Test
    public void fileTypeTest() throws IOException {
        Workbook wb2 = WorkbookFactory.create(Files.newInputStream(Paths.get(PATH + "fileRead.xlsx")));
        Workbook wb = WorkbookFactory.create(Files.newInputStream(Paths.get(PATH + "fileRead.xls")));
        Sheet sheet = wb.getSheetAt(0);
        System.out.println(wb2.getSpreadsheetVersion());
        System.out.println(wb.getSpreadsheetVersion());

        System.out.println(sheet.getRow(0).getCell(2).getStringCellValue());
    }


    @Test
    public void ExcelReaderGetValueTest2() throws FileNotFoundException {
        ExcelReader reader = new ExcelReader(new FileInputStream(PATH + "cellTypeTest.xlsx"));
        for (int i = 0; i < reader.defaultSheet.getPhysicalNumberOfRows(); i++) {
            for (int j = 0; j < reader.defaultSheet.getRow(i).getPhysicalNumberOfCells(); j++) {
                System.out.println("单元的类型：" + reader.defaultSheet.getRow(i).getCell(j).getCellType()
                        + "   值为：" + reader.getValue(i, j));
            }
        }

        System.out.println("===============================================");
        ExcelReader reader2 = new ExcelReader(new FileInputStream(PATH + "fileRead.xlsx"));
        for (int i = 0; i < reader2.defaultSheet.getPhysicalNumberOfRows(); i++) {
            for (int j = 0; j < reader2.defaultSheet.getRow(i).getPhysicalNumberOfCells(); j++) {
                System.out.println("单元的类型：" + reader2.defaultSheet.getRow(i).getCell(j).getCellType()
                        + "   值为：" + reader2.getValue(i, j));
            }
        }

        System.out.println("===============================================");
        ExcelReader reader3 = new ExcelReader(new FileInputStream(PATH + "fileRead.xls"));
        for (int i = 0; i < reader3.defaultSheet.getPhysicalNumberOfRows(); i++) {
            for (int j = 0; j < reader3.defaultSheet.getRow(i).getPhysicalNumberOfCells(); j++) {
                System.out.println("当前行：" + (i + 1) + "当前列：" + (j + 1) +
                        "  单元的类型：" + reader3.defaultSheet.getRow(i).getCell(j).getCellType()
                        + "   值为：" + reader3.getValue(i, j));
            }
        }

        System.out.println(reader3.defaultSheet.getPhysicalNumberOfRows());
        for (int i = 0; i < reader3.defaultSheet.getPhysicalNumberOfRows(); i++) {
            System.out.println("当前行：" + (i + 1) + "   单元格数：" + reader3.defaultSheet.getRow(i).getPhysicalNumberOfCells());
        }

        //System.out.println(reader3.getValue(31, 8));

    }

    @Test
    public void alignmentTest() {
        ExcelWriter writer = new ExcelWriter();

        writer.setValue(0, 0, "qwe").setCellStyle(writer.HorizontalCenter());
        writer.setValue(0, 2, "asd").setCellStyle(writer.VerticalCenter());
        writer.setValue(3, 2, "zxc").setCellStyle(writer.HorizontalVerticalCenter());

        writer.saveFile(PATH + "fileWriter1-align.xlsx");
    }


    @Test
    public void BorderTest() {
        ExcelWriter writer = new ExcelWriter();

        writer.setValue(1, 1, "四个")
                .setCellStyle(writer.setBorderType(
                        BorderStyle.DOUBLE,
                        BorderStyle.NONE,
                        BorderStyle.HAIR,
                        BorderStyle.MEDIUM));

        writer.setValue(2, 2, "三个")
                .setCellStyle(writer.setBorderType(
                        BorderStyle.DASH_DOT_DOT,
                        BorderStyle.SLANTED_DASH_DOT,
                        BorderStyle.DOTTED));

        writer.setValue(3, 3, "两个")
                .setCellStyle(writer.setBorderType(
                        BorderStyle.THICK,
                        BorderStyle.MEDIUM_DASHED));

        writer.setValue(4, 4, "一个")
                .setCellStyle(writer.setBorderType(BorderStyle.THIN));

        writer.saveFile(PATH + "fileWriter2-borderStyle.xlsx");


    }

    @Test
    public void cellExistTest() {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet();
        sheet.createRow(0).createCell(0).setCellValue(123);
        sheet.getRow(0).createCell(0).setCellValue(456);

        System.out.println(sheet.getRow(0));
        System.out.println(sheet.getRow(0).getCell(0));

    }

    @Test
    public void borderColorTest() {
        ExcelWriter writer = new ExcelWriter();

        writer.setValue(1, 1, "颜色").setCellStyle(
                writer.setBorderStyle(writer.setBorderType(
                                BorderStyle.DOUBLE,
                                BorderStyle.NONE,
                                BorderStyle.HAIR,
                                BorderStyle.MEDIUM
                        ),
                        IndexedColors.ORANGE)
        );
        writer.saveFile(PATH + "fileWriter3-borderColor.xlsx");
    }

    @Test
    public void cellMergerTest() {
        ExcelWriter writer = new ExcelWriter();

        writer.setValue(3, 3, "合并");
        writer.cellMerge(3, 5, 3, 5);

        writer.saveFile(PATH + "fileWriter4-cellMerger.xlsx");
    }

    @Test
    public void cellWidthHeightTest() {
        ExcelWriter writer = new ExcelWriter();

        for (int i = 0; i < 4; i++) {
            for (int j = 0; j < 8; j++) {
                writer.setValue(i, j, "瑠璃璃啊" + (i + 1) * (j + 1));
            }
        }

        writer.setColWidth(1, 10);
        System.out.println(writer.sheet.getColumnWidth(1));
        System.out.println(writer.sheet.getColumnWidth(2));
        System.out.println("-----------------------------------");
        System.out.println(writer.sheet.getRow(1).getHeight());
        System.out.println(writer.sheet.getRow(2).getHeight());
        System.out.println(writer.sheet.getRow(2).getHeightInPoints());
        System.out.println(writer.sheet.getRow(2).getZeroHeight());
        System.out.println("=============================");
        writer.sheet.getRow(1).setHeight((short) 400);
        writer.sheet.getRow(2).setHeightInPoints(20);

        System.out.println(writer.sheet.getRow(1).getHeight());
        System.out.println(writer.sheet.getRow(2).getHeight());
        System.out.println(writer.sheet.getRow(2).getHeightInPoints());
        System.out.println(writer.sheet.getRow(2).getZeroHeight());

        writer.setRowHeight(3, 3);

        writer.saveFile(PATH + "fileWriter5-cellWidthHeight.xlsx");

    }

    /**
     * Word文档测试
     */
    @Test
    public void wordReaderDemoTest() throws IOException {
        XWPFDocument doc = new XWPFDocument(new FileInputStream(PATH + "wordReaderDEMO.docx"));

        // 段落
        List<XWPFParagraph> paragraphs = doc.getParagraphs();
        // 表格
        List<XWPFTable> tables = doc.getTables();
        // 图片
        List<XWPFPictureData> allPictures = doc.getAllPictures();
        // 页眉
        List<XWPFHeader> headerList = doc.getHeaderList();
        // 页脚
        List<XWPFFooter> footerList = doc.getFooterList();

        System.out.println(paragraphs.get(5).getParagraphText());
        System.out.println(paragraphs.size());

        System.out.println(tables.get(0).getRow(0).getCell(0).getText());

        System.out.println(paragraphs.get(27).getParagraphText());

    }

    @Test
    public void wordWriterDemoTest2() {
        XWPFDocument document = new XWPFDocument();
        XWPFParagraph paragraph = document.createParagraph();
        XWPFRun run = paragraph.createRun();

        run.setText("Hello, World!");
        run.setFontSize(12);
        run.setBold(true);
        run.setItalic(false);
        run.setColor("0000FF"); // 设置文本颜色

        // 创建第一个段落
        XWPFParagraph paragraph1 = document.createParagraph();
        XWPFRun run1 = paragraph1.createRun();
        run1.setText("This is the first paragraph.");
        run1.setFontSize(12);

        // 创建第二个段落
        XWPFParagraph paragraph2 = document.createParagraph();
        XWPFRun run2 = paragraph2.createRun();
        run2.setText("This is the second paragraph.");
        run2.setFontSize(14);
        run2.setBold(true);
        run2.setText("QQQQQQQQQQ");

        // 将文档保存到文件
        try (FileOutputStream out = new FileOutputStream(PATH + "example.docx")) {
            document.write(out);
        } catch (IOException e) {
            e.printStackTrace();
        }


    }

    @Test
    public void WordWriterTest1() {
        WordWriter wordWriter = new WordWriter();
        XWPFParagraph paragraph = wordWriter.getParagraph();

        paragraph.setAlignment(ParagraphAlignment.CENTER);

        XWPFRun textRun = paragraph.createRun();

        textRun.setText("这是一级标题");
        textRun.setFontFamily("黑体");
        textRun.setFontSize(16);
        textRun.setUnderline(UnderlinePatterns.DOUBLE);
        XWPFRun insertNewRun = paragraph.insertNewRun(0);
        insertNewRun.setText("在段落起始位置插入这段文本");

        XWPFTable table = wordWriter.document.createTable(3, 3);

        table.getRow(1).getCell(1).setText("EXAMPLE OF TABLE");

        XWPFTableCell cell = table.getRow(2).getCell(1);
        XWPFParagraph cp = cell.addParagraph();
        XWPFRun r = cp.createRun();
        r.setText("复杂！！！");

        // 获取单元格段落后设置对齐方式
        XWPFParagraph addParagraph = cell.addParagraph();
        addParagraph.setAlignment(ParagraphAlignment.CENTER);

        wordWriter.saveToFile(PATH + "wordWriter1.docx");

    }

    @Test
    public void WordWriterTest3() {
        WordWriter doc = new WordWriter();
        //创建一个表格，并指定宽度
        XWPFTable table = doc.document.createTable(4, 4);

        XWPFTableCell tableCell = table.getRow(1).getCell(0);
        tableCell.setText("啊啊啊");

        doc.mergeCellsHorizontal(table, 1, 0, 3);

        tableCell.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.RESTART);
        table.setWidth("100%");


        //TableTools.widthTable(table, MiniTableRenderData.WIDTH_A4_FULL, 4);


        //TableStyle ts = new TableStyle();
        //table.setTopBorder(XWPFTable.XWPFBorderType.DOTTED, 1, 2, "#2b2b2b");


        //设置第0行数据
        List<XWPFTableCell> row0 = table.getRow(0).getTableCells();
        row0.get(0).setText("xxxx"); //为第0行第0列设置内容
        row0.get(0).setWidth("200");
        row0.get(1).setText("aaaa");
        row0.get(2).setText("bbbb");
        row0.get(3).setText("cccc");


        doc.saveToFile(PATH + "wordWriter3.docx");


    }

    @Test
    public void WordWriterTest4() {
        WordWriter writer = new WordWriter();

        //createTable(writer.wordFile);

        XWPFTable table = writer.document.createTable(3, 4);
        for (int i = 0; i < 3; i++) {
            for (int j = 0; j < 4; j++) {
                table.getRow(i).getCell(j).setText("第" + (i + 1) + "行第" + (j + 1) + "列");
            }
        }


        table.setWidth("100%");

        CTTblBorders borders = table.getCTTbl().getTblPr().addNewTblBorders();
        CTBorder hBorder = borders.addNewInsideH();
        hBorder.setVal(STBorder.Enum.forString("single"));  // 线条类型
        hBorder.setSz(new BigInteger("1")); // 线条大小
        hBorder.setColor("000000"); // 设置颜色

        CTBorder vBorder = borders.addNewInsideV();
        vBorder.setVal(STBorder.Enum.forString("single"));
        vBorder.setSz(new BigInteger("1"));
        vBorder.setColor("000000");

        CTBorder lBorder = borders.addNewLeft();
        lBorder.setVal(STBorder.Enum.forString("single"));
        lBorder.setSz(new BigInteger("1"));
        lBorder.setColor("000000");

        CTBorder rBorder = borders.addNewRight();
        rBorder.setVal(STBorder.Enum.forString("single"));
        rBorder.setSz(new BigInteger("1"));
        rBorder.setColor("000000");

        CTBorder tBorder = borders.addNewTop();
        tBorder.setVal(STBorder.Enum.forString("single"));
        tBorder.setSz(new BigInteger("1"));
        tBorder.setColor("000000");

        CTBorder bBorder = borders.addNewBottom();
        bBorder.setVal(STBorder.Enum.forString("single"));
        bBorder.setSz(new BigInteger("1"));
        bBorder.setColor("000000");

        table.createRow(); // 增加一行

        writer.saveToFile(PATH + "wordWriter4.docx");

    }

    @Test
    public void WordWriterTest5() {
        WordWriter writer = new WordWriter();

        XWPFTable table = writer.getTable(3, 4);

        for (int i = 0; i < 3; i++) {
            for (int j = 0; j < 4; j++) {
                writer.setTableCellValue(i, j, "第" + (i + 1) + "行第" + (j + 1) + "列");
            }
        }

        //writer.mergeCellsHorizontal(table, 0, 0, 3);
        //writer.mergeCellsVertically(table, 0, 0, 2);

        writer.saveToFile(PATH + "wordWriter5.docx");
    }

    @Test
    public void convertedWorkBook() {
        ExcelWriter.saveToPDF(PATH + "workbook.xls", PATH + "新_电子信息工程2023版指标点.pdf");
    }

}
