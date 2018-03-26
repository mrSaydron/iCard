package ru.mrak.iCard.util;

import ru.mrak.iCard.FProperties;

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

    private static String egexp = "^(ИГУЛ|IGUL)\\.?(\\d{6})\\.?(\\d{3})(-(\\d{2}))?\\s*_?([А-Яа-яA-Za-z]{1,2}[0-9]{0,2})?(_([0-9]{1,2}))?(\\s(.*))?\\.([A-Za-z]{3,5})$";

    private NameParser() {}

    public static NameParser parser(String fileName, FProperties properties) {
        NameParser np = new NameParser();
        Pattern pattern = Pattern.compile(properties.getNameRegexp());
        Matcher matcher = pattern.matcher(fileName);
        if(matcher.find()) {
            np.organizationCod = matcher.group(properties.getOrganizationCod());
            np.characteristic = matcher.group(properties.getCharacteristic());
            np.registrationNumber = matcher.group(properties.getRegistrationNumber());
            np.species = matcher.group(properties.getSpecies());
            np.documentCode = matcher.group(properties.getDocumentCode());
            np.documentVersion = matcher.group(properties.getDocumentVersion());
            np.documentName = matcher.group(properties.getDocumentName());
            np.filenameExtension = matcher.group(properties.getFilenameExtension());
            return np;
        }
        return null;
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
