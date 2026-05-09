from pathlib import Path

CONFIG = {
    # --- Режим работы ---
    "DRY_RUN": False,
    "dry_run_file_email": Path("dry_run_email.html"),
    "dry_run_file_jira": Path("dry_run_jira.txt"),
    # --- Логирование ---
    "log_file": Path("check_hrad_users.log"),
    "log_retention_days": 60,
    # --- Excel ---
    # Полный путь до реестра. Адаптируйте под свою инфраструктуру.
    # Для Windows-share: r"\\server\share\path\Exclusions.xlsx"
    "excel_path": Path(
        r"C:\Reports\InfoSec\Registries\Exclusions.xlsx"
    ),
    "sheet_name": "HR_AD_ExcludeAll",
    "header_row": 2,
    # --- Active Directory ---
    "base_dn": "DC=corp,DC=local",
    # --- Jira ---
    "jira_keyring_service_name": "python_jira_script_excludeall",
    "jira_server": "https://jira.example.com",
    "jira_quarterly_task_jql": (
        "project = 'SOC' "
        "AND issuetype = 'Task' "
        "AND summary ~ 'Periodic check of employeeType attribute' "
        "AND resolution = Unresolved "
        "ORDER BY created DESC"
    ),
    "jira_project_key": "SOC",
    "jira_new_issue_type": "Task",
    "jira_new_issue_summary": (
        "Production: ExcludeAll Report: Обнаружены несоответствия"
    ),
    "jira_link_type": "Связано с",
    # --- Почта ---
    "smtp_server": "smtp.corp.local",
    "from_email": "soc@example.com",
    "to_emails": [
        # "admin@example.com",
        "soc@example.com",
    ],
    "email_subject": "Проверка учётных записей HR/AD ExcludeAll — результаты",
    "smtp_retries": 3,
    "smtp_delay_sec": 5,
    "smtp_timeout_sec": 20,
}
