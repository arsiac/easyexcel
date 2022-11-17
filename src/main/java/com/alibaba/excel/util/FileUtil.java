package com.alibaba.excel.util;

import org.apache.commons.compress.archivers.zip.ZipArchiveEntry;
import org.apache.commons.compress.archivers.zip.ZipFile;
import org.apache.commons.compress.utils.IOUtils;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.util.Enumeration;

/**
 * @author jipengfei
 */
public class FileUtil {

    private static final int BUF = 4096;

    public static void writeFile(File file, InputStream stream) {
        final String filePath = file.getAbsolutePath();
        makeDirs(filePath);
        if (!file.exists()) {
            create(file);
        }

        try (OutputStream os = new FileOutputStream(file)) {
            byte[] data = new byte[1024];
            int length;
            while ((length = stream.read(data)) != -1) {
                os.write(data, 0, length);
            }
            os.flush();
        } catch (FileNotFoundException e) {
            throw new IllegalStateException("FileNotFoundException occurred. ", e);
        } catch (IOException e) {
            throw new IllegalStateException("IOException occurred. ", e);
        }
    }

    public static String getFolderName(String filePath) {
        if (filePath == null || "".equals(filePath)) {
            return filePath;
        }
        final int pos = filePath.lastIndexOf(File.separator);
        return (pos == -1) ? "" : filePath.substring(0, pos);
    }

    /**
     * 文件解压
     */
    public static void unzip(String path, File file) throws IOException {
        try (ZipFile zipFile = new ZipFile(file, StandardCharsets.UTF_8.name())) {
            Enumeration<ZipArchiveEntry> en = zipFile.getEntries();
            ZipArchiveEntry ze;
            while (en.hasMoreElements()) {
                ze = en.nextElement();
                if (ze.getName().contains("../")) {
                    //防止目录穿越
                    throw new IllegalStateException("insecurity zip file!");
                }
                File f = new File(path, ze.getName());
                if (ze.isDirectory()) {
                    mkdirs(f);
                    continue;
                } else {
                    mkdirs(f.getParentFile());
                }

                try (InputStream is = zipFile.getInputStream(ze);
                     OutputStream os = Files.newOutputStream(f.toPath())) {
                    IOUtils.copy(is, os, BUF);
                }
            }
        }
    }

    public static void mkdirs(File file) {
        if (file == null) {
            return;
        }
        if (!file.mkdirs()) {
            throw new IllegalStateException("mkdirs failed: " + file.getAbsolutePath());
        }
    }

    private static void makeDirs(String filePath) {
        String folderName = getFolderName(filePath);
        if (folderName == null || "".equals(folderName)) {
            return;
        }
        File folder = new File(folderName);
        if (!folder.exists() || !folder.isDirectory()) {
            mkdirs(folder);
        }
    }

    public static void deleteFile(String filepath) {
        File file = new File(filepath);
        // 当且仅当此抽象路径名表示的文件存在且 是一个目录时，返回 true
        if (file.isDirectory()) {
            String[] children = file.list();
            if (children != null && children.length > 0) {
                for (String childName : children) {
                    deleteFile(filepath + File.separator + childName);
                }
            }
            delete(file);
        } else {
            delete(file);
        }
    }

    private static void delete(File file) {
        try {
            Files.delete(file.toPath());
        } catch (IOException e) {
            throw new IllegalStateException("delete failed: " + file.getAbsolutePath(), e);
        }
    }

    private static void create(File file) {
        try {
            if (!file.createNewFile()) {
                throw new IllegalStateException("Create file failed: " + file.getAbsolutePath());
            }
        } catch (IOException e) {
            throw new IllegalStateException("Create file failed: " + file.getAbsolutePath(), e);
        }
    }

    private FileUtil() {
    }

}
