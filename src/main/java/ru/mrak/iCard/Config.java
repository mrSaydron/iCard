package ru.mrak.iCard;

import java.nio.charset.Charset;
import java.util.*;

/**Парсит конфигурацию запуска прогрммы и хранит результаты
 * При значении следующих ключей равном true бедет произведено
 *  rename - переименование найденных файлов к нормируемому виду;
 *  delete - удаление исходных фалов;
 *  createWB - создание документа "Информационно-удостовверяющий лист";
 *  assayFilenameExtension - сравнение расширения файлов из рабочей папки со списком допустимых расширений для выбора с
 *  какими файлами работать
 *
 * При неправильно заданной конфигурации на консоль выводится предупреждение и приложение завершается
 */
public class Config {
    private boolean rename;
    private boolean delete;
    private boolean createWB;
    private boolean assayFilenameExtension;

    private static Character[] configChars = {'r', 'd', 'c', 'a'};

    private Config() {}

    /**Метод парсит строку конфигурации
     * строка представляет из себя "-rdca", ключи:
     * r(ename) - переименование файлов в формат IGUL123456789;
     * d(elete) - удаляет исходные файлы;
     * c(reate) - создает информационно-удостоверяющий лист;
     * a(ssay) - работает с файлами имеющими расширение из файла свойств.
     * ключь может повторятся
     * При неправильно заданной конфигурации на консоль выводится предупреждение и приложение завершается
     * @param config строка конфигурации
     * @return объект с ключами конфигурации
     */
    public static Config parsing(String config) {
        if(config == null || config.isEmpty()) return null;
        if(config.length() < 2) return null;
        if(config.charAt(0) != '-') return null;
        config = config.substring(1);
        Set<Character> configSet = new HashSet<>();
        for(char c : config.toCharArray()) {
            configSet.add(c);
        }
        Set<Character> configCharsSet = new HashSet<Character>(Arrays.asList(configChars));
        if(!configCharsSet.containsAll(configSet)) {
            System.out.println("Конфигурация задана неправильно");
            System.exit(0);
        }
        Config conf = new Config();
        if(configSet.contains(configChars[0])) conf.rename = true;
        if(configSet.contains(configChars[1])) conf.delete = true;
        if(configSet.contains(configChars[2])) conf.createWB = true;
        if(configSet.contains(configChars[3])) conf.assayFilenameExtension = true;

        return conf;
    }

    public String createConfig() {
        StringBuilder string = new StringBuilder("-");
        if(rename) string.append('r');
        if(delete) string.append('d');
        if(createWB) string.append('c');
        if(assayFilenameExtension) string.append('a');
        return string.toString();
    }

    public boolean isRename() {
        return rename;
    }

    public void setRename(boolean rename) {
        this.rename = rename;
    }

    public boolean isDelete() {
        return delete;
    }

    public void setDelete(boolean delete) {
        this.delete = delete;
    }

    public boolean isCreateWB() {
        return createWB;
    }

    public void setCreateWB(boolean createWB) {
        this.createWB = createWB;
    }

    public boolean isAssayFilenameExtension() {
        return assayFilenameExtension;
    }

    public void setAssayFilenameExtension(boolean assayFilenameExtension) {
        this.assayFilenameExtension = assayFilenameExtension;
    }

    @Override
    public String toString() {
        return "Config{" +
                "rename=" + rename +
                ", delete=" + delete +
                ", createWB=" + createWB +
                ", assayFilenameExtension=" + assayFilenameExtension +
                '}';
    }

}
