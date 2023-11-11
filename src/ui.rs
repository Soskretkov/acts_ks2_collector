use crate::constants::CONSOLE_LEFT_MARGIN_IN_SPACES;
use console::{Style, Term};
use dialoguer::Input;
use std::io;
use std::path::PathBuf;
use std::thread; // для засыпания на секунду-две
use std::time::Duration; // для засыпания на секунду-две // для очистки консоли перед выводом полезных сообщений

pub fn user_input() -> (PathBuf, String) {
    loop {
        let entered_text = inputting_path();
        let path = PathBuf::from(&entered_text);

        if path.exists() {
            break (path, entered_sheet_name());
        }

        //filter нужен на случай ввода "details"  в кавычках (@ - на случай русской раскладки)
        let keyword = entered_text
            .chars()
            .filter(|ch| *ch != '"' && *ch != '@')
            .collect::<String>()
            .to_lowercase();
        let len_text = keyword.chars().count();

        match keyword {
            x if len_text < 9
                && x.matches([
                    'd', 'e', 't', 'a', 'i', 'l', 's', 'в', 'у', 'е', 'ф', 'ш', 'д', 'ы',
                ])
                .count()
                    > 4 =>
            {
                display_details();
                thread::sleep(Duration::from_secs(2));
                continue;
            }
            _ => continue,
        }
    }
}

pub fn display_first_lines(is_visible: bool) {
    let optional_text = if is_visible {
        "       Введите  \"details\"  для получения подробностей о программе."
    } else {
        ""
    };

    let msg = format!("\n{optional_text}\n");
    display_formatted_text(&msg, None);
}

pub fn display_help() {
    let msg = format!(
        r#"------------------------------------------------------------------------------------------------------------

● Используйте CTRL + V, чтобы вставить скопированный путь к папке или файлу с данными, которые вы хотите собрать.");
● Программа будет собирать данные из файлов Excel по указанному пути, включая вложенные папки.");
● Собираются только файлы с расширением «.xlsm».");
● Полезный совет:
    - переименуйте файл Excel, добавив символ «@», и программа не будет собирать его данные;");
    - переименуйте папку, добавив символ «@», и программа проигнорирует ее содержимое.

------------------------------------------------------------------------------------------------------------"#
    );

    display_formatted_text(&msg, None);
}

pub fn display_formatted_text(text: &str, text_style: Option<Style>) {
    let formatted_text = prepend_spaces_to_non_empty_lines(text);

    match text_style {
        Some(style) => println!("{}", style.apply_to(formatted_text)),
        None => println!("{}", formatted_text),
    }
}

fn display_details() {
    // Очистка прошлых сообщений
    let _ = Term::stdout().clear_screen();
    display_first_lines(false);
    display_help();

    let msg = format!(
        r#"
            Наименование продукта:        «Сборщик данных из актов формы \"КС-2\"»
            Версия продукта:              {}
            Дата основания проекта:       02.06.2022
            Адрес на GitHub.com:          https://github.com/Soskretkov/ks2_etl
            Автор:                        Оскретков Сергей Юрьевич
            Специально для: ООО «Трест Росспецэнергомонтаж»,
            Альтуфьевское шоссе, д. 43, стр. 1, Москва, 127410,
            Cметно-договорное управление.

------------------------------------------------------------------------------------------------------------"#,
        env!("CARGO_PKG_VERSION")
    );
    display_formatted_text(&msg, None);
}

fn inputting_path() -> String {
    display_formatted_text("Введите путь:", None);
    let mut text = String::new();
    io::stdin()
        .read_line(&mut text)
        .expect("Ошибка чтения ввода");

    text = text.trim().to_string();
    text
}

fn entered_sheet_name() -> String {
    let _ = Term::stdout().clear_screen();
    let msg = prepend_spaces_to_non_empty_lines(
        "Подтвердите лист или укажите другой.
    Не имеет значения, используете ли вы прописные или строчные буквы при указании листа.
    Имя листа",
    );
    let entered_sh_name: String = Input::new()
        .with_prompt(msg)
        .with_initial_text("Лист1")
        .interact()
        .expect("Ошибка чтения ввода");

    let _ = Term::stdout().clear_screen();
    entered_sh_name
}

fn prepend_spaces_to_non_empty_lines(text: &str) -> String {
    let spaces = " ".repeat(CONSOLE_LEFT_MARGIN_IN_SPACES);
    text.lines()
        .map(|line| {
            if !line.trim().is_empty() {
                format!("{}{}", spaces, line)
            } else {
                line.to_string()
            }
        })
        .collect::<Vec<String>>()
        .join("\n")
}
