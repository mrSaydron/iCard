package ru.mrak.iCard;

import org.apache.commons.codec.binary.Hex;
import ru.mrak.iCard.util.NameParser;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.*;
import java.nio.file.attribute.BasicFileAttributes;
import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Объект работает с файлами в рабочей директории
 * в зависимости от ключей переданых через объект Config переименовывает фалй к нормированым,
 * удаляет исходные файлы и отфильтровывать рабочие файлы
 * Возвращает список подготовленных данных собраных из обработанных файлов
 */
public class FilesReplace {
    private String source;
    private Path sourcePath;
    private String result;
    private Path resultPath;
    private final List<EntryDoc> entryDocs;

    private boolean rename;
    private boolean delete;
    private boolean createWB;
    private boolean assayFilenameExtension;

    private String[] documentCodesCyr;
    private String[] documentCodesLat;
    private String[] filenameExtensions;
    private FProperties properties;

    private static String filenameExtensionPattern = ".+\\.(\\w{3,4})";

    private FilesReplace(String source, String result) {
        this.source = source;
        this.result = result;
        entryDocs = new ArrayList<>();
    }

    /**
     * Метод выполняет переименовывает, удаляет и собирате информацию в зависимости от переданных данных
     * @param source - директория к исходным файлам
     * @param result - директория для сохранения результатов работы метода
     * @param config - конфигурирует метод
     * @param properties - набор свойств для конфигурирования метода
     * @return - возвращет список параметров созданный из считанных файлов
     */
    public static List<EntryDoc> replace(String source, String result, Config config, FProperties properties) {
        FilesReplace filesReplace = new FilesReplace(source, result);
        filesReplace.rename = config.isRename();
        filesReplace.delete = config.isDelete();
        filesReplace.createWB = config.isCreateWB();
        filesReplace.assayFilenameExtension = config.isAssayFilenameExtension();
        filesReplace.documentCodesCyr = properties.getDocumentCodeCyr();
        filesReplace.documentCodesLat = properties.getDocumentCodeLat();
        filesReplace.filenameExtensions = properties.getFilenameExtensions();
        filesReplace.properties = properties;

        filesReplace.checkSourcePath();
        filesReplace.createResultPath();
        filesReplace.walkFiles();

        return filesReplace.entryDocs;
    }

    /**
     * Проверяет существует ли папка с исходными файлами
     */
    private void checkSourcePath() {
        sourcePath = Paths.get(source);
        if(sourcePath == null || !Files.exists(sourcePath)) {
            if(sourcePath != null)
                System.out.println("Такого пути не существует: " + sourcePath.toAbsolutePath().toString());
            else
                System.out.println("Такого пути не существует: " + source);
            System.exit(0);
        }
        if(!Files.isDirectory(sourcePath)) {
            System.out.println("Не папка: " + sourcePath.toAbsolutePath().toString());
            System.exit(0);
        }
        System.out.println("Папка с исходниками: " + sourcePath.toAbsolutePath());
    }

    /**
     * Создает папку для результатов
     */
    private void createResultPath() {
        resultPath = Paths.get(result);
        if(Files.exists(resultPath) && !Files.isDirectory(resultPath)) {
            System.out.println(resultPath.toAbsolutePath() + " - не папка");
            System.exit(0);
        }
        try {
            if(!Files.exists(resultPath))
                Files.createDirectories(resultPath);
        } catch (IOException e) {
            System.out.println("Не удалось создать папку: " + resultPath.toAbsolutePath().toString());
            System.exit(0);
        }
        System.out.println("Папка с результатом: " + resultPath.toAbsolutePath());
    }

    /**
     * Копипует файл
     * @param file - исходный файл
     * @param newFile - результирующий файл
     */
    private void copyFile(Path file, Path newFile) {
        try {
            Files.copy(file, newFile, StandardCopyOption.REPLACE_EXISTING);
        } catch (IOException e) {
            System.out.println("Не удалось скопировать: " + file.getFileName());
            System.exit(0);
        }
    }

    /**
     * Копирует файл и считает его контролтную сумму
     * @param file - исходный файл
     * @param newFile - результирующий файл
     * @return контрольная сумма исходного файла
     */
    private String copyFileAndComputeMd5(Path file, Path newFile) {
        MessageDigest messageDigest = null;
        try {
            messageDigest = MessageDigest.getInstance("MD5");
        } catch (NoSuchAlgorithmException e) {
            System.out.println("Нет алгоритма шифрования MD5");
            System.exit(0);
        }
        try (InputStream is = Files.newInputStream(file);
             OutputStream os = Files.newOutputStream(newFile)){
            byte[] buffer = new byte[1024];
            int length;
            messageDigest.reset();
            while ((length = is.read(buffer)) > 0) {
                messageDigest.update(buffer, 0, length);
                os.write(buffer, 0, length);
            }
            byte[] digest = messageDigest.digest();
            return Hex.encodeHexString(digest);
        } catch (IOException e) {
            System.out.println("Не удалось скопировать: " + file.getFileName());
            System.exit(0);
        }
        return null;
    }

    /**
     * Считает контрольную сумму файла
     * @param file - файл члоя расчета контрольной суммы
     * @return контрольную сумму файла
     */
    private String computeMd5(Path file) {
        MessageDigest messageDigest = null;
        try {
            messageDigest = MessageDigest.getInstance("MD5");
        } catch (NoSuchAlgorithmException e) {
            System.out.println("Нет алгоритма шифрования MD5");
            System.exit(0);
        }
        try (InputStream is = Files.newInputStream(file)){
            byte[] buffer = new byte[1024];
            int length;
            messageDigest.reset();
            while ((length = is.read(buffer)) > 0) {
                messageDigest.update(buffer, 0, length);
            }
            byte[] digest = messageDigest.digest();
            return Hex.encodeHexString(digest);
        } catch (IOException e) {
            System.out.println("Не удалось скопировать: " + file.getFileName());
            System.exit(0);
        }
        return null;
    }

    /**
     * Обходит папку проверяя файлы на соответсктвие шаблону
     * С файлами прошедшими проверку проводит операции в соответсвии с конфигурацией
     */
    private void walkFiles() {
        System.out.println("Поиск файлов по: " + NameParser.getEgexp());
        System.out.println("\n");
        final Set<Path> renameFiles = new HashSet<>();
        try {
            Files.walkFileTree(sourcePath, new HashSet<>(), 1, new SimpleFileVisitor<Path>() {
                @Override
                public FileVisitResult visitFile(Path file, BasicFileAttributes attrs){
                    if (renameFiles.contains(file)) return FileVisitResult.CONTINUE;
                    String fileName = file.getFileName().toString();
                    //Проерека на расширение
                    if(assayFilenameExtension) {
                        Pattern pattern = Pattern.compile(filenameExtensionPattern);
                        Matcher matcher = pattern.matcher(fileName);
                        if(!matcher.find() || Arrays.asList(filenameExtensions).contains(matcher.group(1))) {
                            return FileVisitResult.CONTINUE;
                        }
                    }
                    //Проверка на соответствие наименованию документов
                    if (!fileName.matches(properties.getNameRegexp())) return FileVisitResult.CONTINUE;
                    System.out.println("Файл: " + fileName);

                    NameParser document = NameParser.parser(fileName, properties);
                    EntryDoc entryDoc = new EntryDoc();

                    //Определение кода документа
                    String documentCodLat = null;
                    String documentCodCyr = null;
                    int index = -1;
                    if(document.getDocumentCode() != null) {
                        if (document.getOrganizationCod().equals("IGUL")) {
                            index = Arrays.asList(documentCodesLat).indexOf(document.getDocumentCode().toLowerCase());
                        } else {
                            index = Arrays.asList(documentCodesCyr).indexOf(document.getDocumentCode().toUpperCase());
                        }
                        if(index >= 0) {
                            documentCodCyr = documentCodesCyr[index];
                            documentCodLat = documentCodesLat[index];
                        } else {
                            documentCodCyr = document.getDocumentCode();
                            documentCodLat = document.getDocumentCode();
                        }

                    }
                    //Обозначение документа
                    entryDoc.setDesignation("ИГУЛ." +
                            document.getCharacteristic() +
                            "." +
                            document.getRegistrationNumber() +
                            (documentCodCyr != null ? (" " + documentCodCyr) : ""));
                    //Наименование документа
                    entryDoc.setName(document.getDocumentName());
                    //Наименование файла
                    if(rename) {
                        entryDoc.setFileName("IGUL" +
                            document.getCharacteristic() +
                            document.getRegistrationNumber() +
                            document.getSpeciesAndDash() +
                            (documentCodLat != null ? ("_" + documentCodLat) : "") +
                            (document.getDocumentVersion() != null ? ("_" + document.getDocumentVersion()) : "") +
                            "." +
                            document.getFilenameExtension());
                    } else {
                        entryDoc.setFileName(fileName);
                    }
                    //Дата редактирования
                    BasicFileAttributes fileAttributes = null;
                    try {
                        fileAttributes = Files.readAttributes(file, BasicFileAttributes.class);
                    } catch (IOException e) {
                        System.out.println("Не удалось считать атрибуты файла: " + file.toString());
                        System.exit(0);
                    }
                    DateFormat dateFormat = new SimpleDateFormat("dd.MM.yyyy");
                    entryDoc.setDate(dateFormat.format(fileAttributes.lastModifiedTime().toMillis()));
                    //Размер файла
                    entryDoc.setSize("" + fileAttributes.size());
                    //Версия
                    entryDoc.setVersion(document.getDocumentVersion() != null ? (document.getDocumentVersion()) : "-");
                    //Номер релиза
                    entryDoc.setReleaseNumber("-");
                    //Копирование и подсчет MD5
                    Path newFile = Paths.get(resultPath + "\\" + entryDoc.getFileName());
                    if(rename) {
                        if(createWB) {
                            if(file.equals(newFile)) {
                                entryDoc.setMd5(computeMd5(file));
                            } else {
                                entryDoc.setMd5(copyFileAndComputeMd5(file, newFile));
                            }
                        }
                        copyFile(file, newFile);
                    } else {
                        if(createWB) {
                            entryDoc.setMd5(computeMd5(file));
                        }
                    }
                    //Удаление файла
                    if(delete) {
                        try {
                            Files.delete(file);
                        } catch (IOException e) {
                            System.out.println("Не удалось удалить файл: " + file.toString());
                            System.exit(0);
                        }
                    }
                    //Вывожк результаты в консоль
                    System.out.println(entryDoc);
                    entryDocs.add(entryDoc);
                    //Сохраняю файл чтобы не использовать второй раз
                    renameFiles.add(newFile);
                    return FileVisitResult.CONTINUE;
                }
            });
        } catch (IOException e) {
            System.out.println("Не удалось скопировать файлы");
            System.exit(0);
        }
    }


}
