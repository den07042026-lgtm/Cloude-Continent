# /// script
# requires-python = ">=3.10"
# dependencies = ["playwright", "openpyxl"]
# ///
"""
gpt_design_prompt.py
════════════════════
Разовый запрос: просим ChatGPT изучить типовые описания автозапчастей на
Ozon/WB (веб-поиск) и самому составить промпт-инструкцию для генерации
описаний. Результат сохраняется в designed_prompt.txt для проверки перед
тем, как зашить его в gpt_ozon500_descriptions.py.

Запуск:
  uv run --with playwright scripts/gpt_design_prompt.py
"""

import sys
import time
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, str(Path(__file__).parent))

from playwright.sync_api import sync_playwright
from gpt_ozon500_descriptions import (
    CHATGPT_URL, CDP_PORT, dismiss_modals, type_message,
    enable_web_search, click_send, wait_for_response, clean_description,
)

META_PROMPT = (
    "Изучи в интернете несколько реальных примеров описаний автозапчастей на "
    "карточках товаров Ozon и Wildberries - разные типы: масляные фильтры, "
    "тормозные колодки, амортизаторы, сайлентблоки, катушки зажигания и т.п., "
    "разные бренды. Посмотри, как устроены хорошие, технически грамотные "
    "описания, которые реально помогают покупателю выбрать нужную деталь и не "
    "ошибиться (а не просто маркетинговый текст).\n\n"
    "На основе этого анализа составь промпт-инструкцию на русском языке, "
    "которую я буду использовать для генерации похожих описаний для своих "
    "товаров (у меня есть только наименование, бренд, артикул и категория "
    "детали для каждого товара - промпт должен предполагать, что дальше эти "
    "данные будут подставляться в шаблон).\n\n"
    "Требования к результату:\n"
    "- объём описания 1200-1500 символов\n"
    "- технически точный текст, без маркетинговых клише и воды\n"
    "- обязательно помогает покупателю не ошибиться при выборе (совместимость, "
    "на что обратить внимание)\n"
    "- пригоден и для Ozon, и для Wildberries\n\n"
    "Ответь ТОЛЬКО текстом самого промпта (который я скопирую и буду "
    "использовать), без вступления и пояснений до или после."
)


def main():
    with sync_playwright() as pw:
        browser = pw.chromium.connect_over_cdp(f"http://127.0.0.1:{CDP_PORT}")
        context = browser.contexts[0] if browser.contexts else browser.new_context()
        page = None
        for p in context.pages:
            if "chatgpt.com" in p.url:
                page = p
                break
        if page is None:
            page = context.pages[0] if context.pages else context.new_page()

        page.goto(CHATGPT_URL, wait_until="domcontentloaded", timeout=30_000)
        time.sleep(3)
        dismiss_modals(page)

        print("Ввожу мета-промпт...")
        type_message(page, META_PROMPT)

        web_search_on = enable_web_search(page)
        print(f"Веб-поиск: {'включен' if web_search_on else 'не найден в UI'}")

        click_send(page)
        time.sleep(2)

        response = wait_for_response(page, timeout_sec=300 if web_search_on else 180)
        result = clean_description(response)

        out_path = Path(__file__).parent.parent / "data" / "analytics" / "top500_ozon" / "designed_prompt.txt"
        out_path.parent.mkdir(parents=True, exist_ok=True)
        out_path.write_text(result, encoding="utf-8")

        print()
        print("=" * 62)
        print("РЕЗУЛЬТАТ (сохранён в", out_path, "):")
        print("=" * 62)
        print(result)


if __name__ == "__main__":
    main()
