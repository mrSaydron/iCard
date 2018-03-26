package ru.mrak.iCard;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Properties;

/**
 * Считывает конфигурацию из папки с приложением и из рабочей папки
 * Конфигурирует:
 * documentCodeCyr - список кодов документа на кирилице
 * documentCodeLat - список кодов документа на латинице
 * filenameExtensions - список расширений файлов с которым работает приложение
 * author - автор информационно-удостоверяющего листа
 * checked - проверяющий информационно-удостоверяющего листа
 * approved - утверждающий информационно-удостоверяющий лист
 */
public class FProperties {
    private String[] documentCodeCyr = {"СБ", "МЭ", "ТУ"};
    private String[] documentCodeLat = {"sb", "me", "tu"};
    private String[] filenameExtensions = {"dwg", "tdd", "xls"};
    private String author = "Разработал";
    private String checked = "";
    private String approved = "";
    private String nameRegexp = "^(ИГУЛ|IGUL)\\.?(\\d{6})\\.?(\\d{3})(-(\\d{2}))?\\s*_?([А-Яа-яA-Za-z]{1,2}[0-9]{0,2})?(_([0-9]{1,2}))?(\\s(.*))?\\.([A-Za-z]{3,5})$";
    private int organizationCod = 1;
    private int characteristic = 2;
    private int registrationNumber = 3;
    private int species = 5;
    private int documentCode = 6;
    private int documentVersion = 8;
    private int documentName = 10;
    private int filenameExtension = 11;

    /**
     * Пытается прочитать конфигурацию сначала из папки с приложением, затем из рабочей папки
     * Если считать не удалось, выводит предупреждение в консоль и возвращает объект с значениями по умолчанию
     */
    public FProperties(String source) {
        String jarPath = FProperties.class.getProtectionDomain().getCodeSource().getLocation().getPath();
        String globalPath = jarPath.substring(1,jarPath.lastIndexOf("/")) + "/properties.ini";
        Path globalProperties = Paths.get(globalPath);
        Properties properties = new Properties();
        if(Files.exists(globalProperties)) {
            try(InputStream stream = Files.newInputStream(globalProperties)) {
                properties.load(stream);
                readProperties(properties);
            } catch (IOException e) {
                System.out.println("Не получилось открыть файл глобальных свойств по пути: " + globalProperties.toAbsolutePath().toString());
            }
        } else {
            System.out.println("Нет файла глобальных свойств по пути: " + globalProperties.toAbsolutePath().toString());
        }
        Path localProperties = Paths.get(source + "\\properties.ini");
        if(Files.exists(localProperties)) {
            try (InputStreamReader stream = new InputStreamReader(new FileInputStream(localProperties.toAbsolutePath().toString()), "UTF-8")) {
                properties.load(stream);
                readProperties(properties);
            } catch (IOException e) {
                System.out.println("Не получилось открыть файл локальных свойств по пути: " + localProperties.toAbsolutePath().toString());
            }
        } else {
            System.out.println("Нет файла локальных свойств по пути: " + localProperties.toAbsolutePath().toString());
        }
    }

    private void readProperties(Properties properties) {
        if(properties.containsKey("documentCodeCyr")) {
            documentCodeCyr = properties.getProperty("documentCodeCyr").split("[,;\\s]");
        }
        if(properties.containsKey("documentCodeLat")) {
            documentCodeLat = properties.getProperty("documentCodeLat").split("[,;\\s]");
        }
        if(properties.containsKey("filenameExtensions")) {
            filenameExtensions = properties.getProperty("filenameExtensions").split("[,;\\s]");
        }
        if(properties.containsKey("author")) {
            author = properties.getProperty("author");
        }
        if(properties.containsKey("checked")) {
            checked = properties.getProperty("checked");
        }
        if(properties.containsKey("approved")) {
            approved = properties.getProperty("approved");
        }
        if(properties.containsKey("nameRegexp")) {
            nameRegexp = properties.getProperty("nameRegexp");
        }
        if(properties.containsKey("organizationCod")) {
            try {
                organizationCod = Integer.parseInt(properties.getProperty("organizationCod"));
            } catch (NumberFormatException e) {
                System.out.println("Свойство organizationCod должно быть числом");
            }
        }
        if(properties.containsKey("characteristic")) {
            try {
                organizationCod = Integer.parseInt(properties.getProperty("characteristic"));
            } catch (NumberFormatException e) {
                System.out.println("Свойство characteristic должно быть числом");
            }
        }
        if(properties.containsKey("registrationNumber")) {
            try {
                organizationCod = Integer.parseInt(properties.getProperty("registrationNumber"));
            } catch (NumberFormatException e) {
                System.out.println("Свойство registrationNumber должно быть числом");
            }
        }
        if(properties.containsKey("species")) {
            try {
                organizationCod = Integer.parseInt(properties.getProperty("species"));
            } catch (NumberFormatException e) {
                System.out.println("Свойство species должно быть числом");
            }
        }
        if(properties.containsKey("documentCode")) {
            try {
                organizationCod = Integer.parseInt(properties.getProperty("documentCode"));
            } catch (NumberFormatException e) {
                System.out.println("Свойство documentCode должно быть числом");
            }
        }
        if(properties.containsKey("documentVersion")) {
            try {
                organizationCod = Integer.parseInt(properties.getProperty("documentVersion"));
            } catch (NumberFormatException e) {
                System.out.println("Свойство documentVersion должно быть числом");
            }
        }
        if(properties.containsKey("documentName")) {
            try {
                organizationCod = Integer.parseInt(properties.getProperty("documentName"));
            } catch (NumberFormatException e) {
                System.out.println("Свойство documentName должно быть числом");
            }
        }
        if(properties.containsKey("filenameExtension")) {
            try {
                organizationCod = Integer.parseInt(properties.getProperty("filenameExtension"));
            } catch (NumberFormatException e) {
                System.out.println("Свойство filenameExtension должно быть числом");
            }
        }
    }

    public String[] getDocumentCodeCyr() {
        return documentCodeCyr;
    }

    public String[] getDocumentCodeLat() {
        return documentCodeLat;
    }

    public String[] getFilenameExtensions() {
        return filenameExtensions;
    }

    public String getAuthor() {
        return author;
    }

    public String getChecked() {
        return checked;
    }

    public String getApproved() {
        return approved;
    }

    public String getNameRegexp() {
        return nameRegexp;
    }

    public int getOrganizationCod() {
        return organizationCod;
    }

    public int getCharacteristic() {
        return characteristic;
    }

    public int getRegistrationNumber() {
        return registrationNumber;
    }

    public int getSpecies() {
        return species;
    }

    public int getDocumentCode() {
        return documentCode;
    }

    public int getDocumentVersion() {
        return documentVersion;
    }

    public int getDocumentName() {
        return documentName;
    }

    public int getFilenameExtension() {
        return filenameExtension;
    }
}
