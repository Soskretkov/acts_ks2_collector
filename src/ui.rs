use console::Term;
use std::io;
use std::thread;
use std::time::Duration;

pub fn session() -> (String, String) {
    show_first_lines();
    show_help();

    loop {
        let path = inputting_path();

        if path.matches(['\\']).count() > 0 {
            break (path, inputting_sheet_name());
        }

        let len_path = path.chars().count();

        match path {
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

        //filter нужен на случай ввода "details"  в кавычках (@ - на случай русской раскладки)
        text = text
            .trim()
            .chars()
            .filter(|ch| *ch != '"' && *ch != '@')
            .collect::<String>()
            // .trim_end_matches('\\')
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
        let _ = Term::stdout().clear_screen();
        println!("Введите имя листа:");
        thread::sleep(Duration::from_secs(1));
        println!("Нет разницы, вводите ли вы «Лист1» или «лист1» - способ указания листа не чувствителен к регистру.");
        let mut temp_sh_name = String::new();

        io::stdin()
            .read_line(&mut temp_sh_name)
            .expect("Ошибка чтения ввода");

        temp_sh_name = temp_sh_name.trim().to_string();

        let len_text = temp_sh_name.chars().count();

        if len_text > 0 {
            return temp_sh_name.to_lowercase();
        }
    }
}

fn show_first_lines() {
    println!("        Введите  \"details\"  для получения подробностей о программе.\n");
}
#[rustfmt::skip]
fn show_help() {
    println!("------------------------------------------------------------------------------------------------------------\n");
    println!("● Используйте CTRL + C, чтобы вставить скопированный путь к папке, из которой необходимо собрать данные;");
    println!("● Программа будет собирать данные из файлов в указанной папке и всех вложенных папках;");
    println!("● Собираются только файлы с расширением «.xlsm»;");
    println!("● Полезный совет: переименуйте файл Excel, добавив символ «@», и программа не будет собирать его данные.");
    println!("\n------------------------------------------------------------------------------------------------------------\n");
}
#[rustfmt::skip]
fn show_details() {
    // Очистка прошлых сообщений
    let _ = Term::stdout().clear_screen();
    println!("\n");
    show_help();

    //println!("------------------------------------------------------------------------------------------------------------\n");
    println!("            Наименование продукта:        «Сборщик данных из актов формы \"КС-2\"»");
    println!("            Адрес проекта на GitHub.com:  https://github.com/Soskretkov/ks2_etl");
    println!("            Дата создания:                01.07.2022");
    println!("            Дата последних изменений:     01.07.2022");
    println!("            Автор:                        Оскретков Сергей Юрьевич\n");
    println!("            Специально для: ООО «Трест Росспецэнергомонтаж»,");
    println!("            Альтуфьевское шоссе, д. 43, стр. 1, Москва, 127410,");
    println!("            Cметно-договорное управление.");
    println!(
        "\n------------------------------------------------------------------------------------------------------------\n"
    );
}
