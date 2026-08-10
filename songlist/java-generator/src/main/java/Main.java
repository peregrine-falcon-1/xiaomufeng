import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.nio.file.*;
import java.util.*;

public class Main {

    // =========================================================
    // Excel 文件
    // =========================================================

    private static final String EXCEL_FILE =
            "C:/Users/14666/Downloads/xiaomufeng.xlsx";


    // =========================================================
    // 背景 SVG
    // =========================================================

    private static final String BACKGROUND_FILE =
            "C:/Users/14666/Downloads/背景.svg";


    // =========================================================
    // 个人高光 SVG 文件夹
    // =========================================================

    private static final String HIGHLIGHT_FOLDER =
            "C:/Users/14666/Downloads/个人高光";


    // =========================================================
    // 网站项目目录
    // =========================================================

    private static final String PROJECT_DIR =
            "C:/Users/14666/Downloads/SongList";


    // =========================================================
    // JSON 输出位置
    // =========================================================

    private static final String JSON_FILE =
            PROJECT_DIR + "/data/songs.json";


    // =========================================================
    // 背景图片输出位置
    // =========================================================

    private static final String BACKGROUND_OUTPUT =
            PROJECT_DIR + "/images/background.svg";


    // =========================================================
    // 个人高光 SVG 输出位置
    // =========================================================

    private static final String HIGHLIGHT_OUTPUT =
            PROJECT_DIR + "/images/highlights";


    // =========================================================
    // Excel Sheet
    // =========================================================

    private static final String[] SHEET_NAMES = {
            "流行",
            "民谣",
            "粤语",
            "偏古风",
            "个人高光"
    };


    // =========================================================
    // 主程序
    // =========================================================

    public static void main(String[] args) {

        System.out.println();
        System.out.println("================================");
        System.out.println("       歌单网站文件生成器");
        System.out.println("================================");


        // -----------------------------------------------------
        // 创建网站目录
        // -----------------------------------------------------

        createDirectory(
                PROJECT_DIR + "/data"
        );

        createDirectory(
                PROJECT_DIR + "/images"
        );

        createDirectory(
                HIGHLIGHT_OUTPUT
        );


        // -----------------------------------------------------
        // 读取 Excel
        // -----------------------------------------------------

        Map<String, List<String>> songLists =
                readExcel();


        // -----------------------------------------------------
        // 生成 songs.json
        // -----------------------------------------------------

        writeJson(songLists);


        // -----------------------------------------------------
        // 复制背景 SVG
        // -----------------------------------------------------

        copyBackground();


        // -----------------------------------------------------
        // 复制个人高光 SVG
        // -----------------------------------------------------

        copyHighlightImages();


        // -----------------------------------------------------
        // 输出结果
        // -----------------------------------------------------

        System.out.println();
        System.out.println("================================");
        System.out.println("              完成");
        System.out.println("================================");


        for (
                Map.Entry<String, List<String>> entry
                        : songLists.entrySet()
        ) {

            System.out.println(
                    entry.getKey()
                            + "："
                            + entry.getValue().size()
                            + " 首"
            );
        }


        System.out.println();

        System.out.println(
                "JSON："
                        + JSON_FILE
        );

        System.out.println(
                "背景："
                        + BACKGROUND_OUTPUT
        );

        System.out.println(
                "个人高光："
                        + HIGHLIGHT_OUTPUT
        );


        System.out.println(
                "================================"
        );
    }


    // =========================================================
    // 读取 Excel
    // =========================================================

    private static Map<String, List<String>> readExcel() {

        Map<String, List<String>> result =
                new LinkedHashMap<>();


        File excelFile =
                new File(EXCEL_FILE);


        // -----------------------------------------------------
        // 检查 Excel 是否存在
        // -----------------------------------------------------

        if (!excelFile.exists()) {

            System.err.println();

            System.err.println(
                    "找不到 Excel："
            );

            System.err.println(
                    EXCEL_FILE
            );

            return result;
        }


        try (
                FileInputStream input =
                        new FileInputStream(excelFile);

                Workbook workbook =
                        new XSSFWorkbook(input)
        ) {

            DataFormatter formatter =
                    new DataFormatter();


            // -------------------------------------------------
            // 按指定顺序读取五个 Sheet
            // -------------------------------------------------

            for (String sheetName : SHEET_NAMES) {

                Sheet sheet =
                        workbook.getSheet(sheetName);


                // -------------------------------------------------
                // Sheet 不存在
                // -------------------------------------------------

                if (sheet == null) {

                    System.out.println(
                            "警告：没有找到 Sheet："
                                    + sheetName
                    );

                    result.put(
                            sheetName,
                            new ArrayList<>()
                    );

                    continue;
                }


                List<String> songs =
                        new ArrayList<>();


                // -------------------------------------------------
                // 读取第一列
                // -------------------------------------------------

                for (Row row : sheet) {

                    Cell cell =
                            row.getCell(
                                    0,
                                    Row.MissingCellPolicy
                                            .RETURN_BLANK_AS_NULL
                            );


                    if (cell == null) {
                        continue;
                    }


                    String songName =
                            formatter
                                    .formatCellValue(cell)
                                    .trim();


                    // 忽略空行

                    if (songName.isEmpty()) {
                        continue;
                    }


                    songs.add(songName);
                }


                result.put(
                        sheetName,
                        songs
                );
            }


        } catch (Exception e) {

            System.err.println();

            System.err.println(
                    "读取 Excel 失败！"
            );

            e.printStackTrace();
        }


        return result;
    }


    // =========================================================
    // 生成 songs.json
    // =========================================================

    private static void writeJson(
            Map<String, List<String>> songLists
    ) {

        File jsonFile =
                new File(JSON_FILE);


        try {

            File parent =
                    jsonFile.getParentFile();


            if (parent != null) {
                parent.mkdirs();
            }


            try (
                    BufferedWriter writer =
                            new BufferedWriter(
                                    new OutputStreamWriter(
                                            new FileOutputStream(
                                                    jsonFile
                                            ),
                                            StandardCharsets.UTF_8
                                    )
                            )
            ) {

                writer.write("{\n");


                int categoryIndex = 0;


                for (
                        Map.Entry<String, List<String>> entry
                                : songLists.entrySet()
                ) {

                    String category =
                            entry.getKey();

                    List<String> songs =
                            entry.getValue();


                    writer.write("  \"");

                    writer.write(
                            escapeJson(category)
                    );

                    writer.write("\": [\n");


                    // -------------------------------------------------
                    // 写入歌曲
                    // -------------------------------------------------

                    for (
                            int i = 0;
                            i < songs.size();
                            i++
                    ) {

                        writer.write("    \"");

                        writer.write(
                                escapeJson(
                                        songs.get(i)
                                )
                        );

                        writer.write("\"");


                        if (
                                i < songs.size() - 1
                        ) {

                            writer.write(",");
                        }


                        writer.write("\n");
                    }


                    writer.write("  ]");


                    if (
                            categoryIndex
                                    < songLists.size() - 1
                    ) {

                        writer.write(",");
                    }


                    writer.write("\n");


                    categoryIndex++;
                }


                writer.write("}\n");
            }


            System.out.println(
                    "songs.json 生成成功！"
            );


        } catch (IOException e) {

            System.err.println(
                    "生成 JSON 失败！"
            );

            e.printStackTrace();
        }
    }


    // =========================================================
    // 复制背景 SVG
    // =========================================================

    private static void copyBackground() {

        File source =
                new File(BACKGROUND_FILE);

        File destination =
                new File(BACKGROUND_OUTPUT);


        if (!source.exists()) {

            System.err.println();

            System.err.println(
                    "找不到背景 SVG："
            );

            System.err.println(
                    BACKGROUND_FILE
            );

            return;
        }


        try {

            destination.getParentFile()
                    .mkdirs();


            Files.copy(
                    source.toPath(),
                    destination.toPath(),
                    StandardCopyOption.REPLACE_EXISTING
            );


            System.out.println(
                    "背景 SVG 复制成功！"
            );


        } catch (IOException e) {

            System.err.println(
                    "复制背景 SVG 失败！"
            );

            e.printStackTrace();
        }
    }


    // =========================================================
    // 复制个人高光 SVG
    // =========================================================

    private static void copyHighlightImages() {

        File sourceFolder =
                new File(HIGHLIGHT_FOLDER);


        File outputFolder =
                new File(HIGHLIGHT_OUTPUT);


        // -----------------------------------------------------
        // 检查文件夹
        // -----------------------------------------------------

        if (!sourceFolder.exists()) {

            System.err.println();

            System.err.println(
                    "找不到个人高光文件夹："
            );

            System.err.println(
                    HIGHLIGHT_FOLDER
            );

            return;
        }


        if (!sourceFolder.isDirectory()) {

            System.err.println();

            System.err.println(
                    "个人高光路径不是文件夹："
            );

            System.err.println(
                    HIGHLIGHT_FOLDER
            );

            return;
        }


        // -----------------------------------------------------
        // 创建输出文件夹
        // -----------------------------------------------------

        outputFolder.mkdirs();


        // -----------------------------------------------------
        // 找到所有 SVG
        // -----------------------------------------------------

        File[] svgFiles =
                sourceFolder.listFiles(
                        file ->
                                file.isFile()
                                        && file.getName()
                                        .toLowerCase()
                                        .endsWith(".svg")
                );


        if (
                svgFiles == null
                        || svgFiles.length == 0
        ) {

            System.err.println();

            System.err.println(
                    "个人高光文件夹中没有 SVG！"
            );

            return;
        }


        // -----------------------------------------------------
        // 按文件名排序
        // -----------------------------------------------------

        Arrays.sort(
                svgFiles,
                Comparator.comparing(
                        File::getName,
                        String.CASE_INSENSITIVE_ORDER
                )
        );


        int count = 0;


        // -----------------------------------------------------
        // 复制 SVG
        // -----------------------------------------------------

        for (File svg : svgFiles) {

            try {

                File destination =
                        new File(
                                outputFolder,
                                svg.getName()
                        );


                Files.copy(
                        svg.toPath(),
                        destination.toPath(),
                        StandardCopyOption
                                .REPLACE_EXISTING
                );


                count++;


            } catch (IOException e) {

                System.err.println(
                        "复制失败："
                                + svg.getName()
                );

                e.printStackTrace();
            }
        }


        System.out.println(
                "个人高光 SVG："
                        + count
                        + " 张"
        );
    }


    // =========================================================
    // 创建目录
    // =========================================================

    private static void createDirectory(
            String path
    ) {

        File directory =
                new File(path);


        if (!directory.exists()) {

            if (directory.mkdirs()) {

                System.out.println(
                        "创建目录："
                                + path
                );
            }
        }
    }


    // =========================================================
    // JSON 特殊字符处理
    // =========================================================

    private static String escapeJson(
            String text
    ) {

        return text
                .replace("\\", "\\\\")
                .replace("\"", "\\\"")
                .replace("\b", "\\b")
                .replace("\f", "\\f")
                .replace("\n", "\\n")
                .replace("\r", "\\r")
                .replace("\t", "\\t");
    }
}
