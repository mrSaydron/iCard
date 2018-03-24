package ru.mrak.iCard.util;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Класс парсит именя файла в:
 * organizationCod - код организации
 * characteristic - характеристику
 * registrationNumber - регистрационный номер
 * species - исполнение
 * documentCode - код документа
 * documentName - наименование
 * filenameExtension - расширение файла
 *
 * egexp - регулярное выражение для парсинга строки
 */
public class NameParser {
    private String organizationCod;
    private String characteristic;
    private String registrationNumber;
    private String species;
    private String documentCode;
    private String documentVersion;
    private String documentName;
    private String filenameExtension;

    private static String egexp = "^(ИГУЛ|IGUL)\\.?(\\d{6})\\.?(\\d{3})(-(\\d{2}))?\\s*_?([А-Яа-яA-Za-z]{1,2}[0-9]{0,2})?(_([0-9]{1,2}))?\\s+?(.*)\\.([A-Za-z]{3})$";

    private NameParser() {}

    public static NameParser parser(String fileName, String egexp) {
        NameParser np = new NameParser();
        Pattern pattern = Pattern.compile(egexp);
        Matcher matcher = pattern.matcher(fileName);
        if(matcher.find()) {
            np.organizationCod = matcher.group(1);
            np.characteristic = matcher.group(2);
            np.registrationNumber = matcher.group(3);
            np.species = matcher.group(5);
            np.documentCode = matcher.group(6);
            np.documentVersion = matcher.group(8);
            np.documentName = matcher.group(9);
            np.filenameExtension = matcher.group(10);
            return np;
        }
        return null;
    }

    public static NameParser parser(String fileName) {
        return parser(fileName, egexp);
    }

    public String getOrganizationCod() {
        return organizationCod;
    }

    public String getCharacteristic() {
        return characteristic;
    }

    public String getRegistrationNumber() {
        return registrationNumber;
    }

    public String getSpecies() {
        return species;
    }

    public String getSpeciesAndDash() {
        if(species != null && !species.isEmpty())
                return species;
        return "";
    }

    public String getDocumentCode() {
        return documentCode;
    }

    public String getDocumentName() {
        return (documentName != null) ? documentName : "";
    }

    public String getFilenameExtension() {
        return filenameExtension;
    }

    public static String getEgexp() {
        return egexp;
    }

    public String getDocumentVersion() {
        return documentVersion;
    }

    public static void setEgexp(String egexp) {
        NameParser.egexp = egexp;
    }

    @Override
    public String toString() {
        return  "organizationCod=" + organizationCod + '\n' +
                "characteristic=" + characteristic + '\n' +
                "registrationNumber=" + registrationNumber + '\n' +
                "species=" + species + '\n' +
                "documentCode=" + documentCode + '\n' +
                "documentName=" + documentName + '\n' +
                "filenameExtension='" + filenameExtension + '\n';
    }
}
