import os
import re
import time
import random
from datetime import datetime
from urllib.parse import quote

import pandas as pd
from bs4 import BeautifulSoup
from tqdm import tqdm

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager


# =========================
# НАСТРОЙКИ
# =========================

AREA_ID = 113
PAGES_PER_KEYWORD = 3

SAVE_EVERY = 10

OUTPUT_FILE = "hh_big_creative_vacancies_html.xlsx"
CHECKPOINT_FILE = "hh_big_creative_checkpoint_html.xlsx"
LINKS_FILE = "hh_big_creative_links_html.xlsx"
LOG_FILE = "hh_big_parser_html_log.txt"


KEYWORD_GROUPS = {
    "business_creative": [
        "графический дизайнер",
        "бренд-дизайнер",
        "дизайнер",
        "ux/ui дизайнер",
        "motion designer",
        "арт-директор",
        "креативный директор",
        "копирайтер",
        "креативный копирайтер",
        "редактор",
        "smm",
        "smm-специалист",
        "контент-менеджер",
        "контент-креатор",
        "креативный продюсер",
        "маркетолог",
        "pr-менеджер"
    ],

    "artistic_creative": [
        "художник",
        "иллюстратор",
        "живописец",
        "музыкант",
        "композитор",
        "аранжировщик",
        "звукорежиссёр",
        "саунд-дизайнер",
        "вокалист",
        "актёр",
        "режиссёр",
        "сценарист",
        "драматург",
        "хореограф",
        "танцор",
        "фотограф",
        "видеооператор",
        "монтажёр",
        "аниматор",
        "3d artist",
        "3d художник",
        "game artist",
        "concept artist",
        "художник-постановщик",
        "декоратор",
        "костюмер",
        "гример"
    ],

    "hybrid_creative": [
        "архитектор",
        "ландшафтный дизайнер",
        "дизайнер интерьера",
        "дизайнер одежды",
        "промышленный дизайнер",
        "game designer",
        "геймдизайнер",
        "продюсер",
        "театральный продюсер",
        "кино-продюсер",
        "event-менеджер",
        "организатор мероприятий"
    ]
}


# =========================
# ЛОГ
# =========================

def log(message):
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    text = f"[{now}] {message}"
    print(text)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(text + "\n")


def clean_text(text):
    if not text:
        return ""
    return " ".join(str(text).split())


# =========================
# SELENIUM
# =========================

options = Options()
options.add_argument("--start-maximized")
options.add_argument("--disable-blink-features=AutomationControlled")

driver = webdriver.Chrome(
    service=Service(ChromeDriverManager().install()),
    options=options
)


# =========================
# ПАРСИНГ HTML
# =========================

def extract_vacancy_links(html):
    soup = BeautifulSoup(html, "lxml")
    links = set()

    for a in soup.find_all("a", href=True):
        href = a["href"]

        if "/vacancy/" in href:
            match = re.search(r"/vacancy/(\d+)", href)
            if match:
                vacancy_id = match.group(1)
                url = f"https://hh.ru/vacancy/{vacancy_id}"
                links.add((vacancy_id, url))

    return list(links)


def parse_vacancy_page(html):
    soup = BeautifulSoup(html, "lxml")

    h1 = soup.find("h1")
    title = clean_text(h1.get_text()) if h1 else ""

    company_tag = soup.find(attrs={"data-qa": "vacancy-company-name"})
    company = clean_text(company_tag.get_text()) if company_tag else ""

    desc_tag = soup.find(attrs={"data-qa": "vacancy-description"})
    if desc_tag:
        description = clean_text(desc_tag.get_text(" "))
    else:
        description = clean_text(soup.get_text(" "))

    salary_tag = soup.find(attrs={"data-qa": "vacancy-salary"})
    salary = clean_text(salary_tag.get_text()) if salary_tag else ""

    area_tag = soup.find(attrs={"data-qa": "vacancy-view-raw-address"})
    area = clean_text(area_tag.get_text()) if area_tag else ""

    exp_tag = soup.find(attrs={"data-qa": "vacancy-experience"})
    experience = clean_text(exp_tag.get_text()) if exp_tag else ""

    schedule_tag = soup.find(attrs={"data-qa": "vacancy-view-employment-mode"})
    schedule = clean_text(schedule_tag.get_text()) if schedule_tag else ""

    return {
        "name": title,
        "company": company,
        "salary_text": salary,
        "area_text": area,
        "experience_text": experience,
        "schedule_text": schedule,
        "description": description
    }


def maybe_wait_for_human_check():
    """
    Исправленная версия.
    Теперь код НЕ реагирует на обычные слова "проверка", "подтвердите", "телефон" внутри вакансии.
    Он ждёт только если реально видит признаки captcha-страницы.
    """
    html = driver.page_source.lower()
    current_url = driver.current_url.lower()

    captcha_signals = [
        "account/signup/captcha",
        "captcha?backurl",
        "/captcha",
        "hcaptcha",
        "smartcaptcha",
        "g-recaptcha",
        "captcha-container",
        "captcha__",
        "проверка безопасности",
        "подтвердите, что вы не робот",
        "докажите, что вы не робот",
        "я не робот"
    ]

    is_real_captcha = (
        "captcha" in current_url
        or any(signal in html for signal in captcha_signals)
    )

    if is_real_captcha:
        log("Похоже, появилась настоящая капча. Пройди её вручную в браузере. Жду 90 секунд.")
        time.sleep(90)


# =========================
# СБОР ССЫЛОК
# =========================

def collect_links():
    if os.path.exists(LINKS_FILE):
        log(f"Найден файл ссылок: {LINKS_FILE}")
        links_df = pd.read_excel(LINKS_FILE)
        links_df["vacancy_id"] = links_df["vacancy_id"].astype(str)
        log(f"Загружено ссылок: {len(links_df)}")
        return links_df

    all_links = []

    for segment, keywords in KEYWORD_GROUPS.items():
        for keyword in keywords:
            log(f"Ищем: {keyword} / сегмент: {segment}")

            for page in range(PAGES_PER_KEYWORD):
                search_url = (
                    f"https://hh.ru/search/vacancy?"
                    f"text={quote(keyword)}"
                    f"&area={AREA_ID}"
                    f"&items_on_page=50"
                    f"&search_field=name"
                    f"&page={page}"
                )

                log(f"Открываем: {search_url}")
                driver.get(search_url)

                time.sleep(4 + random.random() * 3)
                maybe_wait_for_human_check()

                html = driver.page_source
                links = extract_vacancy_links(html)

                log(f"{keyword}, страница {page}: найдено ссылок {len(links)}")

                for vacancy_id, url in links:
                    all_links.append({
                        "creative_segment": segment,
                        "query_keyword": keyword,
                        "vacancy_id": str(vacancy_id),
                        "url": url
                    })

                time.sleep(1.5 + random.random() * 2)

            links_df = pd.DataFrame(all_links)

            if not links_df.empty:
                links_df = links_df.drop_duplicates(subset="vacancy_id")
                links_df.to_excel(LINKS_FILE, index=False)
                log(f"Промежуточно сохранено ссылок: {len(links_df)}")

    links_df = pd.DataFrame(all_links)

    if links_df.empty:
        log("Ссылки не собраны.")
        return links_df

    links_df = links_df.drop_duplicates(subset="vacancy_id").reset_index(drop=True)
    links_df.to_excel(LINKS_FILE, index=False)

    log(f"Итого уникальных ссылок: {len(links_df)}")
    return links_df


# =========================
# CHECKPOINT
# =========================

def load_checkpoint():
    if os.path.exists(CHECKPOINT_FILE):
        log(f"Найден checkpoint: {CHECKPOINT_FILE}")
        old_df = pd.read_excel(CHECKPOINT_FILE)

        if "vacancy_id" in old_df.columns:
            old_df["vacancy_id"] = old_df["vacancy_id"].astype(str)
            done_ids = set(old_df["vacancy_id"].tolist())
        else:
            done_ids = set()

        log(f"Уже обработано вакансий: {len(done_ids)}")
        return old_df, done_ids

    log("Checkpoint не найден. Начинаем обработку с нуля.")
    return pd.DataFrame(), set()


def save_checkpoint(records):
    temp_df = pd.DataFrame(records)

    if not temp_df.empty:
        temp_df["vacancy_id"] = temp_df["vacancy_id"].astype(str)
        temp_df = temp_df.drop_duplicates(subset="vacancy_id").reset_index(drop=True)
        temp_df.to_excel(CHECKPOINT_FILE, index=False)
        temp_df.to_excel(OUTPUT_FILE, index=False)
        log(f"Сохранено: {len(temp_df)} вакансий")

    return temp_df


# =========================
# СКАЧИВАНИЕ ВАКАНСИЙ
# =========================

def collect_vacancy_pages(links_df):
    old_df, done_ids = load_checkpoint()

    records = []
    if not old_df.empty:
        records = old_df.to_dict("records")

    links_df["vacancy_id"] = links_df["vacancy_id"].astype(str)
    to_process = links_df[~links_df["vacancy_id"].isin(done_ids)].copy()

    log(f"Всего ссылок: {len(links_df)}")
    log(f"Уже было обработано: {len(done_ids)}")
    log(f"Осталось обработать: {len(to_process)}")

    counter = 0

    try:
        for _, row in tqdm(to_process.iterrows(), total=len(to_process)):
            vacancy_id = str(row["vacancy_id"])
            url = row["url"]

            try:
                driver.get(url)

                time.sleep(2.5 + random.random() * 2)
                maybe_wait_for_human_check()

                html = driver.page_source
                parsed = parse_vacancy_page(html)

                records.append({
                    "creative_segment": row.get("creative_segment"),
                    "query_keyword": row.get("query_keyword"),
                    "vacancy_id": vacancy_id,
                    "url": url,
                    **parsed
                })

                counter += 1

                if counter % SAVE_EVERY == 0:
                    save_checkpoint(records)

                time.sleep(1.5 + random.random() * 1.5)

            except Exception as e:
                log(f"Ошибка при обработке {vacancy_id}: {e}")
                save_checkpoint(records)
                time.sleep(10)

    except KeyboardInterrupt:
        log("Остановлено вручную. Сохраняю checkpoint...")
        save_checkpoint(records)
        raise

    final_df = save_checkpoint(records)

    log(f"Готово. Итоговый файл: {OUTPUT_FILE}")
    log(f"Размер итоговой таблицы: {final_df.shape}")

    return final_df


# =========================
# MAIN
# =========================

if __name__ == "__main__":
    log("=== ЗАПУСК БОЛЬШОГО HTML-ПАРСЕРА HH ===")
    log(f"PAGES_PER_KEYWORD = {PAGES_PER_KEYWORD}")

    try:
        links_df = collect_links()

        if links_df.empty:
            log("Нет ссылок для обработки.")
        else:
            collect_vacancy_pages(links_df)

    finally:
        driver.quit()
        log("=== ЗАВЕРШЕНО ===")