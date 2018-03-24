package ru.mrak.iCard;

/**
 * Хранит параметры документа
 * designation - обозначение документа
 * name - наименование документа
 * md5 - контрольная сумма
 * fileName - имя файла
 * size - размер файла
 * date - дата последнего редактирования
 * version - версия
 * releaseNumber - номер релиза
 * author - автор
 */
public class EntryDoc {
    private String designation;
    private String name;
    private String md5;
    private String fileName;
    private String size;
    private String date;
    private String version;
    private String releaseNumber;
    private String author;

    public String getDesignation() {
        return designation;
    }

    public void setDesignation(String designation) {
        this.designation = designation;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getMd5() {
        return md5;
    }

    public void setMd5(String md5) {
        this.md5 = md5;
    }

    public String getFileName() {
        return fileName;
    }

    public void setFileName(String fileName) {
        this.fileName = fileName;
    }

    public String getSize() {
        return size;
    }

    public void setSize(String size) {
        this.size = size;
    }

    public String getDate() {
        return date;
    }

    public void setDate(String date) {
        this.date = date;
    }

    public String getVersion() {
        return version;
    }

    public void setVersion(String version) {
        this.version = version;
    }

    public String getReleaseNumber() {
        return releaseNumber;
    }

    public void setReleaseNumber(String releaseNumber) {
        this.releaseNumber = releaseNumber;
    }

    public String getAuthor() {
        return author;
    }

    public void setAuthor(String author) {
        this.author = author;
    }

    @Override
    public String toString() {
        return  "Обозначение: " + (designation != null ? designation : "") + '\n' +
                "Наименование: " + (name != null ? name : "")+ '\n' +
                "MD5: " + (md5 != null ? md5 : "")+ '\n' +
                "Новое имя файла: " + (fileName != null ? fileName : "")+ '\n' +
                "Размер файла: " + (size != null ? size : "")+ '\n' +
                "Дата последнего редактирования: " + (date != null ? date : "")+ '\n' +
                "Версия: " + (version != null ? version : "")+ '\n' +
                "Номер релиза: " + (releaseNumber != null ? releaseNumber : "")+ '\n' +
                "Автор: " + (author != null ? author : "")+ '\n' +
                "-----------------------\n";
    }
}
