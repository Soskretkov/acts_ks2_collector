use std::io;
use std::thread;
use std::time::Duration;

pub fn session() -> (String, String) {
    show_first_lines();

    loop {
        let path = inputting_path();

        if path.matches(['\\']).count() > 0 {
            break (path, inputting_sheet_name());
        }

        let len_path = path.chars().count();

        match path {
            x if len_path < 6
                && x.matches(['h', 'e', 'l', 'p', 'р', 'у', 'д', 'з']).count() > 2 =>
            {
                show_help();
                thread::sleep(Duration::from_secs(1));
                continue;
            }
            x if len_path < 9
                && x.matches([
                    'd', 'e', 't', 'a', 'i', 'l', 's', 'в', 'у', 'е', 'ф', 'ш', 'д', 'ы',
                ])
                .count()
                    > 4 =>
            {
                show_details();
                thread::sleep(Duration::from_secs(2));
                continue;
            }
            _ => continue,
        }
    }
}

fn inputting_path() -> String {
    loop {
        println!("Введите путь:");
        let mut text = String::new();
        io::stdin()
            .read_line(&mut text)
            .expect("Ошибка чтения ввода");

        //filter нужен на случай help в кавычках (@ - на случай русской раскладки)
        text = text
            .trim()
            .chars()
            .filter(|ch| *ch != '"' && *ch != '@')
            .collect::<String>()
            .to_lowercase();

        let len_text = text.chars().count();

        if len_text < 3 {
            continue;
        }

        break text;
    }
}

fn inputting_sheet_name() -> String {
    loop {
        println!("Введите имя листа:");
        let mut temp_sh_name = String::new();

        io::stdin()
            .read_line(&mut temp_sh_name)
            .expect("Ошибка чтения ввода");

        temp_sh_name = temp_sh_name.trim().to_owned();

        let len_text = temp_sh_name.chars().count();

        if len_text > 0 {
            return temp_sh_name;
        }
    }
}

fn show_first_lines() {
    println!("      \"help\"      получение подробностей о работе с программой.");
    println!("      \"details\"   получение подробностей о программе.\n");
}
fn show_help() {
    print!("{esc}c", esc = 27 as char);
    // std::process::Command::new("clear").status();
    // std::process::Command::new("cls").status().unwrap();
    show_first_lines();
    println!("------------------------------------------------------------------------------------------------------------\n");
    println!("● Используйте CTRL + C, чтобы вставить скопированный путь к папке, из которой необходимо собрать данные;");
    println!("● Программа будет собирать данные из файлов в указанной папке и всех вложенных папках;");
    println!("● Собираются только файлы с расширением «.xlsm»;");
    println!("● Способ указания имени листа не чувствителен к регистру - нет разницы, вводите ли вы «Лист1» или «лист1»;");
    println!("● Полезный совет: переименуйте файл Excel, добавив символ «@», и программа не будет собирать его данные.");
    println!(
        "\n------------------------------------------------------------------------------------------------------------\n"
    );
}
fn show_details() {
    print!("{esc}c", esc = 27 as char);
    show_first_lines();
    println!("------------------------------------------------------------------------------------------------------------\n");
    println!("            Наименование продукта:        «Сборщик данных из актов формы \"КС-2\"»");
    println!("            Наименование на GitHub.com:   «acts_ks2_collector»");
    println!("            Адрес прокта на GitHub.com:   https://github.com/Soskretkov/acts_ks2_collector");
    println!("            Дата создания:                10.06.2022");
    println!("            Дата последних изменений:     10.06.2022");
    println!("            Автор:                        Оскретков Сергей Юрьевич\n\n");
    println!("            Создано специально для ООО «Трест Росспецэнергомонтаж»,");
    println!("            Альтуфьевское шоссе, д. 43, стр. 1, Москва, 127410,");
    println!("            Cметно-договорное управление.");
    println!(
        "\n------------------------------------------------------------------------------------------------------------\n"
    );
}
