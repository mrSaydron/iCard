package ru.mrak.iCard;

import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;

/**
 *Приложение предначначено для анализа файлов и составления на их основе информационно-удостоверяющего листа
 * дополнительно может пееименовать файлы в соотвествии с нормами и удалить исходные файлы
 * Работает в консольном редиме. На вход требует конфигурацию и папки для исходнях файлов и результатов
 */
public class Main {
    public static void main(String[] args) {
        if(args == null || args.length == 0) {
            System.out.println("использование: icard -config [result path] [source path]\n" +
                    "\n" +
                    "приложение переименовывает и создает информационно-удостоверяющий лист из файлов с именами вида: \"ИГУЛ.123456.789-01 СБ Наименование\" или \"IGUL123456789-01_sb\"\n" +
                    "\n" +
                    "config - составлчется из следующих ключей\n" +
                    "   r(ename) - переименование файлов в формат IGUL123456789\n" +
                    "   d(elete) - удаляет исходные файлы\n" +
                    "   c(reate) - создает информационно-удостоверяющий лист\n" +
                    "   a(ssay) - работает с файлами имеющими расширение из файла свойств\n" +
                    "\n" +
                    "например конфигурация -rc переименовает файлы и создает информационно-удостоверяющий лист в независимости от расширения файлов\n" +
                    "\n" +
                    "source path - папка с исходными файлами\n" +
                    "result path - папка для результатов\n" +
                    "\n" +
                    "если result path нет, то исходные файлы и результаты будут сохранены в текущую папку\n" +
                    "если source path нет, то исходные файлы будут взяты из текущей папки");
            System.exit(0);
        }
        String source = "";
        String result = "";
        Config config = Config.parsing(args[0]);
        if(config == null) {
            System.out.println("Конфигурация задана неправильно");
            System.exit(0);
        }
        if(args.length >= 2) result = args[1];
        if(args.length >= 3) source = args[2];

        FProperties properties = new FProperties();
        List<EntryDoc> entryDocs = FilesReplace.replace(toAbsolutePat(source), toAbsolutePat(result), config, properties);
        if(config.isCreateWB())
            WorkBook.writeBook("Уд. лист", toAbsolutePat(result) + "\\УЛ.xls", entryDocs, properties);
        System.out.println("Выполнено");
    }

    private static String toAbsolutePat(String path) {
        if(!path.equals("")) {
            if(!path.contains(":")) {
                Path workingPath = Paths.get("");
                return workingPath.toAbsolutePath().toString() + "\\" + path;
            }
        } else {
            Path workingPath = Paths.get("");
            return workingPath.toAbsolutePath().toString();
        }
        return path;
    }
}
