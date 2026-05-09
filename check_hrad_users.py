import argparse
import logging
import smtplib
import sys
import time
from datetime import datetime, timedelta
from email.mime.text import MIMEText
from pathlib import Path
from typing import Any

import keyring
import pandas as pd
from jira import JIRA
from pyad import adquery

try:
    from hrad_config import CONFIG
except ImportError:
    print("КРИТИЧЕСКАЯ ОШИБКА: не найден файл hrad_config.py в папке скрипта!")
    sys.exit(1)
# --- Константы Excel ---
COL_ACCOUNT = "учетная запись"
COL_STATUS = "статус"
COL_EXPIRES = "срок до"
VAL_STATUS_WITHDRAWN = "изъято"
VAL_STATUS_PERMANENT = "бессрочно"
REQUIRED_COLS = [COL_ACCOUNT, COL_STATUS]
# --- Константы AD ---
ATTR_AD_SAM = "sAMAccountName"
ATTR_AD_TYPE = "employeeType"
VAL_AD_TYPE = "ExcludeAll"
ATTR_AD_UAC = "userAccountControl"
UAC_DISABLED = 2  # Бит "Account Disabled" в userAccountControl


# === ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ===
def _clean_old_logs(file_path: Path, days_to_keep: int) -> None:
    """
    Ротация логов внутри одного файла. Удаляет строки старше N дней.
    ВАЖНО: Ожидает, что первые 19 символов строки — это дата
        (YYYY-MM-DD HH:MM:SS).
    """
    if not file_path.exists():
        return
    temp_path = file_path.with_name(f"{file_path.name}.tmp")
    cutoff_date = datetime.now() - timedelta(days=days_to_keep)
    try:
        with (
            file_path.open("r", encoding="utf-8") as f_in,
            temp_path.open("w", encoding="utf-8") as f_out,
        ):
            for line in f_in:
                if len(line) < 19:
                    f_out.write(line)
                    continue
                try:
                    date_part = line[:19]
                    log_date = datetime.strptime(
                        date_part, "%Y-%m-%d %H:%M:%S"
                    )
                    if log_date > cutoff_date:
                        f_out.write(line)
                except ValueError:
                    f_out.write(line)
        temp_path.replace(file_path)
    except Exception as e:
        if temp_path.exists():
            temp_path.unlink()
        print(
            f"[WARNING] Не удалось очистить старые логи: {e}", file=sys.stderr
        )


def _setup_logging(log_file: Path) -> None:
    """
    Настраивает логирование:
    1. В файл (с добавлением в конец).
    2. В консоль (стандартный вывод).
    Предварительно запускает очистку старых логов.
    """
    _clean_old_logs(log_file, CONFIG["log_retention_days"])
    formatter = logging.Formatter("%(asctime)s %(levelname)s: %(message)s")
    root_logger = logging.getLogger()
    root_logger.setLevel(logging.INFO)
    root_logger.handlers = []
    file_handler = logging.FileHandler(log_file, mode="a", encoding="utf-8")
    file_handler.setFormatter(formatter)
    root_logger.addHandler(file_handler)
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)
    root_logger.addHandler(console_handler)


def _get_jira_token() -> str | None:
    """Загружает токен из Keyring."""
    try:
        service = CONFIG["jira_keyring_service_name"]
        return keyring.get_password(service, "jira_api_token")
    except Exception as e:
        logging.warning(f"Ошибка Keyring: {e}")
        return None


def _has_discrepancies(res: dict[str, Any]) -> bool:
    """Проверяет, есть ли расхождения в результатах анализа."""
    return bool(
        res["missing_in_registry"]
        or res["disabled_in_ad"]
        or res["missing_in_ad"]
        or res["withdrawn_in_excel"]
        or res["expired_list"]
    )


def _jira_table(title: str, rows: list[str]) -> str:
    """Генерирует Jira Wiki таблицу."""
    if not rows:
        return ""
    lines = [f"||{title}||"]
    lines.extend(f"|{user}|" for user in rows)
    lines.append("")
    return "\n".join(lines)


def _html_table(title: str, rows: list[str]) -> str:
    """Генерирует HTML-таблицу."""
    if not rows:
        return ""
    lines = [f"<h4>{title}</h4><table><tr><th>Учетная запись</th></tr>"]
    lines.extend(f"<tr><td>{user}</td></tr>" for user in rows)
    lines.append("</table>")
    return "".join(lines)


# === ЗАГРУЗКА ДАННЫХ ===
def load_excel_data(path: Path, sheet: str, header: int) -> pd.DataFrame:
    """Загружает Excel, проверяет столбцы и нормализует данные."""
    df = pd.read_excel(path, sheet_name=sheet, header=header)
    df.columns = [str(col).strip().lower() for col in df.columns]
    missing = [col for col in REQUIRED_COLS if col not in df.columns]
    if missing:
        raise ValueError(f"В реестре нет обязательных столбцов: {missing}")
    for col in [COL_ACCOUNT, COL_STATUS]:
        df[col] = df[col].astype(str).str.lower().fillna("")
    logging.info(f"Excel загружен: {path}")
    return df


def load_ad_users_info(base_dn: str) -> dict[str, bool]:
    """
    Выгружает логины и статус блокировки.
    Возвращает словарь: {'login': is_disabled},
    где True — учетка отключена, False — активна.
    """
    query = adquery.ADQuery()
    query.execute_query(
        attributes=[ATTR_AD_SAM, ATTR_AD_UAC],
        where_clause=f"{ATTR_AD_TYPE}='{VAL_AD_TYPE}'",
        base_dn=base_dn,
    )
    result = {}
    for row in query.get_results():
        login = row[ATTR_AD_SAM].strip().lower()
        uac = row.get(ATTR_AD_UAC, 0)
        is_disabled = bool(uac & UAC_DISABLED)
        result[login] = is_disabled
    return result


# === АНАЛИЗ ===
def get_expired_users(df_active: pd.DataFrame) -> list[str]:
    """
    Возвращает список пользователей с просроченными правами.
    Учитывает:
    1. Статус или дату со значением 'Бессрочно'.
    2. Пустые даты у временных сотрудников.
    """
    overdue = []
    today = datetime.now().date()
    if COL_EXPIRES not in df_active.columns:
        logging.warning(
            f"Отсутствует столбец '{COL_EXPIRES}' "
            f"для проверки просроченных дат."
        )
        return overdue
    for _, row in df_active.iterrows():
        date_val = row.get(COL_EXPIRES)
        account = row.get(COL_ACCOUNT, "не указано")
        status = str(row.get(COL_STATUS, "")).lower().strip()
        # Игнорируем "Бессрочно" в статусе или в дате
        if VAL_STATUS_PERMANENT in status:
            continue
        if (
            isinstance(date_val, str)
            and VAL_STATUS_PERMANENT in date_val.strip().lower()
        ):
            continue
        # Пустые значения → ошибка "Дата не указана"
        if pd.isna(date_val) or (
            isinstance(date_val, str) and date_val.strip() == ""
        ):
            overdue.append(
                f"{account} (Статус '{status}', но не указана дата окончания!)"
            )
            continue
        # Преобразуем в дату:
        # — если Excel хранит как datetime, берём напрямую (без str → parse)
        # — если строка, парсим с dayfirst=True для русского формата ДД.ММ.ГГГГ
        if isinstance(date_val, (datetime, pd.Timestamp)):
            date_obj = date_val.date()
        else:
            end_date_dt = pd.to_datetime(
                str(date_val).strip(), errors="coerce", dayfirst=True
            )
            if pd.isna(end_date_dt):
                overdue.append(f"{account} (Некорректная дата: '{date_val}')")
                continue
            date_obj = end_date_dt.date()
        # Проверка просрочки
        if date_obj < today:
            overdue.append(
                f"{account} (Просрочено до {date_obj.strftime('%d.%m.%Y')})"
            )
    return overdue


def analyze_discrepancies(
    df: pd.DataFrame, ad_info: dict[str, bool]
) -> dict[str, Any]:
    """
    Сравнивает списки пользователей.
    1. Находит расхождения между Excel и AD.
    2. Выявляет отключенные учетные записи (Disabled) в AD.
    3. Формирует полную статистику для шапки отчета.
    """
    excel_all = set(df[COL_ACCOUNT].dropna().str.strip())
    df_active = df[df[COL_STATUS] != VAL_STATUS_WITHDRAWN]
    excel_active = set(df_active[COL_ACCOUNT].dropna())
    ad_all = set(ad_info.keys())
    disabled = [user for user, is_disabled in ad_info.items() if is_disabled]
    return {
        "missing_in_registry": sorted(ad_all - excel_all),
        "disabled_in_ad": sorted(disabled),
        "missing_in_ad": sorted(excel_active - ad_all),
        "withdrawn_in_excel": sorted(
            ad_all.intersection(excel_all - excel_active)
        ),
        "expired_list": get_expired_users(df_active),
        "stats": {
            "total_excel": len(df),
            "total_excel_active": len(excel_active),
            "total_ad": len(ad_all),
            "disabled_ad": len(disabled),
        },
    }


# === ГЕНЕРАЦИЯ ОТЧЁТОВ ===
def generate_jira_text(res: dict[str, Any]) -> str:
    """
    Генерирует текстовое описание (Wiki Markup) для задачи JIRA.
    Включает статистику и таблицы с расхождениями.
    """
    lines = [
        "Итоги проверки:",
        f"Всего записей в реестре: {res['stats']['total_excel']}",
        (
            f"Всего активных записей пользователей в реестре: "
            f"{res['stats']['total_excel_active']}"
        ),
        (
            f"Всего активных пользователей в Active Directory:"
            f" {res['stats']['total_ad']}"
        ),
        f"Из них отключено: {res['stats']['disabled_ad']}",
        "",
    ]
    lines.append(
        _jira_table(
            "В AD присутствуют, но отсутствуют в реестре",
            res["missing_in_registry"],
        )
    )
    lines.append(
        _jira_table(
            "Учетные записи отключены в AD, но имеют атрибут",
            res["disabled_in_ad"],
        )
    )
    lines.append(
        _jira_table(
            "В реестре присутствуют (без 'изъято'), но отсутствуют в AD",
            res["missing_in_ad"],
        )
    )
    lines.append(
        _jira_table(
            "В AD присутствуют, но в реестре имеют только статус 'изъято'",
            res["withdrawn_in_excel"],
        ),
    )
    if res["expired_list"]:
        lines.append("h3. Просроченные разрешения (Срок до):")
        lines.append("||Пользователь||Статус / срок||")
        for item in res["expired_list"]:
            if "(" in item:
                user, detail = item.split(" (", 1)
                detail = "(" + detail
            else:
                user, detail = item, ""
            lines.append(f"|{user}|{detail}|")
        lines.append("")
    if not _has_discrepancies(res):
        lines.append(
            "{panel:bgColor=#E3FCEF|borderColor=#16A085}"
            "{color:green}Несоответствий не выявлено.{color}\n{panel}"
        )
        lines.append("Все учётные записи соответствуют данным AD.")
    return "\n".join(lines)


def generate_html_email(res: dict[str, Any]) -> str:
    """
    Генерирует HTML-код письма.
    Включает CSS-стили, статистику и таблицы.
    """
    lines: list[str] = [
        "<html><head>",
        "<meta charset='utf-8'/>",
        "<style>",
        "table {border-collapse: collapse; width: auto;} ",
        "th, td {border: 1px solid #ddd; padding: 8px;} ",
        "th {background-color: #f2f2f2;}",
        "</style></head><body>",
        (
            "<h3>Итоги проверки:</h3>"
            "<p>Всего записей в реестре: "
            f"<b>{res['stats']['total_excel']}</b><br>"
            "Всего активных записей пользователей в реестре: "
            f"<b>{res['stats']['total_excel_active']}</b><br>"
            "Всего активных пользователей в Active Directory: "
            f"<b>{res['stats']['total_ad']}</b><br>"
            f"Из них отключено: <b>{res['stats']['disabled_ad']}</b></p>"
        ),
    ]
    lines.append(
        _html_table(
            "В AD присутствуют, но нет в реестре", res["missing_in_registry"]
        )
    )
    lines.append(
        _html_table(
            "Учетные записи отключены в AD, но имеют атрибут",
            res["disabled_in_ad"],
        )
    )
    lines.append(
        _html_table(
            "В реестре присутствуют, но нет в AD", res["missing_in_ad"]
        )
    )
    lines.append(
        _html_table(
            "В AD присутствуют, но статус 'изъято'", res["withdrawn_in_excel"]
        )
    )
    if res["expired_list"]:
        expired_lines = [
            "<h3 style='color:red'>Просроченные разрешения:</h3>"
            "<table><tr><th>Пользователь</th><th>Статус / срок</th></tr>"
        ]
        for item in res["expired_list"]:
            if "(" in item:
                user, detail = item.split(" (", 1)
                detail = "(" + detail
            else:
                user, detail = item, ""
            expired_lines.append(f"<tr><td>{user}</td><td>{detail}</td></tr>")
        expired_lines.append("</table>")
        lines.extend(expired_lines)
    if not _has_discrepancies(res):
        lines.append(
            "<div style='background:#E3FCEF; padding:15px; "
            "border:1px solid #16A085;'>"
            "<h3 style='color:green; margin:0;'>"
            "Несоответствий не выявлено.</h3>"
            "<p>Все учётные записи соответствуют данным AD.</p></div>"
        )
    lines.append("</body></html>")
    return "".join(lines)


# === ОТПРАВКА ===
def create_jira_report(token: str, summary: str, description: str) -> None:
    """
    Ищет квартальную задачу, создает новую задачу с отчетом и связывает их.
    """
    if CONFIG["DRY_RUN"]:
        filename = CONFIG["dry_run_file_jira"]
        logging.warning(
            f"[DRY_RUN] Задача '{summary}' НЕ создана (тестовый режим)."
        )
        try:
            with filename.open("w", encoding="utf-8") as f:
                f.write(f"Тема: {summary}\n")
                f.write("=" * 40 + "\n")
                f.write(description)
            logging.info(f"   -> Текст задачи сохранен в файл '{filename}'")
        except Exception:
            pass
        return
    try:
        jira = JIRA(
            server=CONFIG["jira_server"],
            options={
                "headers": {"Authorization": f"Bearer {token}"},
                "server_info": False,
            },
        )
    except Exception as e:
        logging.error(f"Не удалось подключиться к Jira: {e}")
        return
    try:
        issues = jira.search_issues(
            CONFIG["jira_quarterly_task_jql"], maxResults=1
        )
    except Exception as e:
        logging.error(f"Ошибка при поиске квартальной задачи: {e}")
        return
    if not issues:
        logging.error("Квартальная задача не найдена! Отчет не будет привязан")
        return
    parent = issues[0]
    logging.info(f"Родитель: {parent.key}")
    try:
        new_issue = jira.create_issue(
            fields={
                "project": {"key": CONFIG["jira_project_key"]},
                "summary": summary,
                "description": description,
                "issuetype": {"name": CONFIG["jira_new_issue_type"]},
            }
        )
        logging.info(f"Успешно создана новая задача: {new_issue.key}")
    except Exception as e:
        logging.error(f"Не удалось создать задачу в Jira: {e}")
        return
    try:
        jira.create_issue_link(
            type=CONFIG["jira_link_type"],
            inwardIssue=new_issue.key,
            outwardIssue=parent.key,
        )
        logging.info("Задачи успешно связаны.")
    except Exception as e:
        logging.error(f"Не удалось связать задачи: {e}")


def send_email(subject: str, body: str, to_emails: list[str]) -> None:
    """
    Отправляет итоговый отчёт по почте.
    Делает несколько попыток при ошибках соединения.
    """
    if CONFIG["DRY_RUN"]:
        filename = CONFIG["dry_run_file_email"]
        logging.warning(
            f"[DRY_RUN] Письмо '{subject}' НЕ отправлено (тестовый режим)"
        )
        try:
            with filename.open("w", encoding="utf-8") as f:
                f.write(f"<!-- To: {to_emails} -->\n")
                f.write(f"<!-- Subject: {subject} -->\n")
                f.write(body)
            logging.info(f"   -> Сохранено в '{filename}'")
        except Exception:
            pass
        return
    msg = MIMEText(body, "html", "utf-8")
    msg["Subject"] = subject
    msg["From"] = CONFIG["from_email"]
    msg["To"] = ", ".join(to_emails)
    retries = CONFIG["smtp_retries"]
    for attempt in range(1, retries + 1):
        try:
            with smtplib.SMTP(
                CONFIG["smtp_server"], timeout=CONFIG["smtp_timeout_sec"]
            ) as server:
                server.sendmail(
                    CONFIG["from_email"], to_emails, msg.as_string()
                )
            logging.info(f"Письмо отправлено: {subject} (попытка {attempt})")
            return
        except Exception as e:
            logging.error(
                f"Ошибка при отправке письма "
                f"(попытка {attempt}/{retries}): {e}"
            )
            if attempt < retries:
                time.sleep(CONFIG["smtp_delay_sec"])
            else:
                logging.error(
                    "Все попытки отправки письма неудачны. Остановка."
                )


# === ТОЧКА ВХОДА ===
def main() -> None:
    """Главная функция — запуск пайплайна проверки."""
    # 1. Настройка логов
    _setup_logging(CONFIG["log_file"])
    logging.info("=== Старт скрипта ===")
    try:
        # --- ГЛОБАЛЬНЫЙ БЛОК ЗАЩИТЫ ---
        # Если ошибка произойдёт здесь — это БАГ КОДА
        # В этом случае сработает logging.exception() с полным traceback

        # 2. Аргументы командной строки
        parser = argparse.ArgumentParser()
        parser.add_argument(
            "--dry-run", action="store_true", help="Включить тестовый режим"
        )
        args = parser.parse_args()
        if args.dry_run:
            CONFIG["DRY_RUN"] = True
            logging.warning("!!! ЗАПУЩЕН РЕЖИМ DRY_RUN (ТЕСТОВЫЙ РЕЖИМ)")

        # 3. Авторизация Jira
        token = _get_jira_token()
        if not token:
            logging.warning("Токен Jira не загружен (Keyring пуст).")
            logging.warning(
                "Убедитесь, что вы запустили скрипт set_jira_auth.py"
            )
            logging.warning("Интеграция с Jira будет пропущена.")

        # 4. Загрузка Excel
        try:
            df = load_excel_data(
                CONFIG["excel_path"],
                CONFIG["sheet_name"],
                CONFIG["header_row"],
            )
        except Exception as e:
            # Инфраструктурная ошибка (файл недоступен)
            logging.error(f"Критическая ошибка при чтении Excel: {e}")
            send_email(
                "ОШИБКА: Сбой скрипта HR/AD ExcludeAll (Excel)",
                f"<h3>Не удалось прочитать Excel</h3><p>{e}</p>",
                CONFIG.get("to_emails", []),
            )
            return

        # 5. Загрузка Active Directory
        try:
            ad_info = load_ad_users_info(CONFIG["base_dn"])
            if not ad_info:
                logging.warning("ВНИМАНИЕ: Из AD получено 0 пользователей!")
                return
            logging.info(f"Из AD загружено {len(ad_info)} пользователей.")
        except Exception as e:
            # Инфраструктурная ошибка AD
            logging.error(f"Критическая ошибка при запросе в AD: {e}")
            send_email(
                "ОШИБКА: Сбой скрипта HR/AD ExcludeAll (AD)",
                f"<h3>Ошибка AD</h3><p>{e}</p>",
                CONFIG.get("to_emails", []),
            )
            return

        # 6. Анализ расхождений
        logging.info("Начало сверки данных...")
        res = analyze_discrepancies(df, ad_info)
        has_issues = _has_discrepancies(res)
        logging.info(f"Сверка завершена. Есть расхождения: {has_issues}")

        # 7. Создание задачи в Jira
        if token and has_issues:
            logging.info("Отправка в Jira...")
            date_str = datetime.now().strftime("%d.%m.%Y")
            summary = f"{CONFIG['jira_new_issue_summary']} (от {date_str})"
            create_jira_report(token, summary, generate_jira_text(res))
        elif not token:
            logging.warning("Пропуск Jira (Нет токена)")
        else:
            logging.info("Пропуск Jira (Нет расхождений).")

        # 8. Отправка email-отчёта
        logging.info("Отправка письма...")
        html = generate_html_email(res)
        subj = CONFIG["email_subject"]
        if has_issues:
            subj += ": Обнаружены несоответствия"
        else:
            subj += ": Несоответствий не выявлено"
        send_email(subj, html, CONFIG.get("to_emails", []))
    except Exception:
        # --- ГЛОБАЛЬНАЯ ЛОВУШКА (КРИТИЧЕСКИЕ БАГИ КОДА) ---
        # Сюда попадаем только при ВНУТРЕННЕЙ ошибке (баг, тип данных)
        # Записываем ПОЛНЫЙ traceback для отладки программистом
        logging.exception(
            "КРИТИЧЕСКИЙ СБОЙ: Скрипт аварийно завершился "
            "из-за внутренней ошибки!"
        )
        try:
            send_email(
                "КРИТИЧЕСКИЙ СБОЙ: Сбой скрипта HR/AD ExcludeAll",
                "<h3>Скрипт упал с внутренней ошибкой</h3>"
                "<p>Проверьте логи на сервере для деталей.</p>",
                CONFIG.get("to_emails", []),
            )
        except Exception:
            logging.error("Не удалось отправить SOS-письмо.")
    # Финальное сообщение (всегда выполняется, даже после except)
    logging.info("=== Работа скрипта завершена ===")


if __name__ == "__main__":
    main()
