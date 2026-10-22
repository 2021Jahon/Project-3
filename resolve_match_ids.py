"""
🎯 FLASHSCORE SCORE RESOLVER v2.6 (Excel Formatting Edition)

Собирает ТОЛЬКО:
match_id, Liga, home_team, away_team, home_score, away_score,
ht_home_score, ht_away_score, status_text, match_date

Дополнительно:
  ✅ Очистка Liga от " - Тур X" и "Новое окно"
  ✅ Красивое форматирование Excel (ширина колонок, заморозка шапки, выравнивание)
"""

import asyncio
import logging
import os
import random
import re
import sys
from datetime import datetime
from logging.handlers import RotatingFileHandler
from typing import Any, Optional
from zoneinfo import ZoneInfo

import pandas as pd
from playwright.async_api import async_playwright, BrowserContext
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

# ============================================================================
# КОНФИГУРАЦИЯ
# ============================================================================

INPUT_EXCEL = "matches.xlsx"
OUTPUT_SUFFIX = "_scores"
SHEET_NAME = 0
MATCH_ID_COLUMN = None

LIMIT = None            # Для теста поставь 5, для всех - None
SAVE_EVERY = 20
MAX_CONCURRENT = 3
CHUNK_SIZE = 10

BASE_DOMAIN = "https://www.flashscorekz.com"
DEFAULT_API_DOMAIN = "https://46.flashscore.ninja"
DEFAULT_FEED_SIGN = 'SW9D1eZo'
USER_AGENT = (
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
    'AppleWebKit/537.36 (KHTML, like Gecko) '
    'Chrome/120.0.0.0 Safari/537.36'
)
TIMEZONE = "Asia/Almaty"

# Ширина колонок в Excel (в символах)
COLUMN_WIDTHS = {
    'match_id':       12,
    'Liga':           35,
    'home_team':      28,
    'away_team':      28,
    'home_score':     11,
    'away_score':     11,
    'ht_home_score':  13,
    'ht_away_score':  13,
    'status_text':    12,
    'match_date':     18,
}

# ============================================================================
# ЛОГИРОВАНИЕ
# ============================================================================

handler = RotatingFileHandler('resolver.log', maxBytes=5 * 1024 * 1024, backupCount=3, encoding='utf-8')
handler.setFormatter(logging.Formatter('%(asctime)s | %(levelname)-7s | %(message)s', '%H:%M:%S'))
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)
logger.addHandler(handler)
logger.addHandler(logging.StreamHandler())

STATUS_MAP = {
    '0': 'Не начат', '1': 'Не начат', '2': 'LIVE', '3': 'Завершен',
    '4': 'Отложен', '5': 'Отменен', '6': 'LIVE', '7': 'LIVE',
    '31': 'Завершен',
}

SEP_ROW, SEP_CELL, SEP_KV, SEP_KV_ALT = '~', '¬', '÷', '\uf037'

# ============================================================================
# ФУНКЦИЯ СОХРАНЕНИЯ EXCEL С ФОРМАТИРОВАНИЕМ
# ============================================================================

def save_with_widths(df: pd.DataFrame, path: str, columns: list):
    """Сохраняет DataFrame в Excel и настраивает ширину, шапку и заморозку."""
    try:
        # Записываем только нужные колонки
        df[columns].to_excel(path, index=False)

        wb = load_workbook(path)
        ws = wb.active

        # 1. Ширина колонок
        for i, col in enumerate(columns, start=1):
            letter = get_column_letter(i)
            ws.column_dimensions[letter].width = COLUMN_WIDTHS.get(col, 15)

        # 2. Жирная шапка + по центру + перенос текста
        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

        # 3. Заморозка первой строки (шапка всегда видна при скролле)
        ws.freeze_panes = "A2"

        # 4. Выравнивание счетов по центру
        score_cols = {'home_score', 'away_score', 'ht_home_score', 'ht_away_score'}
        for i, col in enumerate(columns, start=1):
            if col in score_cols:
                letter = get_column_letter(i)
                for cell in ws[letter][1:]:
                    cell.alignment = Alignment(horizontal='center')

        wb.save(path)
    except Exception as e:
        logger.warning(f"⚠️ Ошибка форматирования Excel: {e}")

# ============================================================================
# КЛАСС РЕЗОЛВЕРА
# ============================================================================

class MatchResolver:
    def __init__(self):
        self.feed_sign = DEFAULT_FEED_SIGN
        self.api_domain = DEFAULT_API_DOMAIN
        self.timezone = ZoneInfo(TIMEZONE)
        self.base_domain = BASE_DOMAIN

    async def bootstrap(self, browser) -> BrowserContext:
        context = await browser.new_context(user_agent=USER_AGENT, locale='ru-RU')
        page = await context.new_page()
        try:
            await page.goto(f"{BASE_DOMAIN}/football/", wait_until='domcontentloaded', timeout=60000)
            await asyncio.sleep(2)
            env = await page.evaluate('''() => {
                try {
                    const env = window.environment;
                    if (!env) return null;
                    return {
                        feed_sign: env.feed_sign || (env.config && env.config.feed_sign) || null,
                        api_domain: env.feed_resolver?.local?.[0]?.url || null
                    };
                } catch(e) { return null; }
            }''')
            if env:
                self.feed_sign = env.get('feed_sign') or DEFAULT_FEED_SIGN
                self.api_domain = env.get('api_domain') or DEFAULT_API_DOMAIN
            logger.info(f"🌐 API: {self.api_domain} | Feed: {self.feed_sign}")
            return context
        except Exception as e:
            await context.close()
            raise RuntimeError(f"Bootstrap failed: {e}")

    def _extract_kv_dict(self, text: str) -> dict[str, str]:
        kv = {}
        for row in text.split(SEP_ROW):
            for cell in row.split(SEP_CELL):
                if SEP_KV in cell:
                    k, _, v = cell.partition(SEP_KV)
                    kv[k.strip()] = v
                elif SEP_KV_ALT in cell:
                    k, _, v = cell.partition(SEP_KV_ALT)
                    kv[k.strip()] = v
        return kv

    async def fetch_one(self, context: BrowserContext, match_id: str, base_headers: dict, sem: asyncio.Semaphore) -> dict[str, Any]:
        async with sem:
            headers = dict(base_headers)
            headers['Referer'] = f"{self.base_domain}/match/{match_id}/"

            result = {
                "match_id": match_id, "Liga": "", "home_team": "", "away_team": "",
                "home_score": "", "away_score": "", "ht_home_score": "",
                "ht_away_score": "", "status_text": "", "match_date": ""
            }

            # 1. API dc_1_ (Дата и статус)
            status_code = None
            try:
                resp = await context.request.get(f"{self.api_domain}/46/x/feed/dc_1_{match_id}", headers=headers, timeout=10000)
                if resp.status == 200:
                    text = await resp.text()
                    if len(text) > 50:
                        kv = self._extract_kv_dict(text)
                        if kv.get('DC'):
                            try:
                                ts = int(kv['DC'])
                                result['match_date'] = datetime.fromtimestamp(ts, self.timezone).strftime('%Y.%m.%d %H:%M')
                            except: pass
                        if kv.get('DA'): status_code = kv['DA']
            except: pass

            # 2. API df_sui_1_ (Первый тайм)
            try:
                resp = await context.request.get(f"{self.api_domain}/46/x/feed/df_sui_1_{match_id}", headers=headers, timeout=10000)
                if resp.status == 200:
                    text = await resp.text()
                    if len(text) > 50:
                        for row in text.split(SEP_ROW):
                            kv = self._extract_kv_dict(row)
                            if kv.get('AC') in ('1-й тайм', '1st Half') and 'IG' in kv and 'IH' in kv:
                                result['ht_home_score'] = kv['IG']
                                result['ht_away_score'] = kv['IH']
                                break
            except: pass

            # 3. DOM (FT, Команды, Лига, Дата фолбэк)
            page = None
            try:
                page = await context.new_page()
                await page.goto(f"{self.base_domain}/match/{match_id}/", wait_until='domcontentloaded', timeout=20000)
                await asyncio.sleep(1.5)

                dom_data = await page.evaluate('''() => {
                    const homeEl = document.querySelector('.duelParticipant__home .participant__participantName');
                    const awayEl = document.querySelector('.duelParticipant__away .participant__participantName');

                    // --- ЛИГА (Хлебные крошки + Title) ---
                    let league = "";
                    const breadcrumbs = Array.from(document.querySelectorAll('.breadcrumb__link, .breadcrumb a, [class*="breadcrumb"] a'));
                    if (breadcrumbs.length >= 2) {
                        const country = breadcrumbs[1] ? breadcrumbs[1].textContent.trim() : "";
                        const tournament = breadcrumbs[breadcrumbs.length - 1].textContent.trim();
                        league = country ? `${country} ${tournament}` : tournament;
                    }

                    if (!league || league.length < 3) {
                        const parts = document.title.split('|');
                        if (parts.length >= 2) league = parts[1].trim();
                    }

                    // --- ОЧИСТКА ЛИГИ (Убираем " - Тур X" и "Новое окно") ---
                    league = league.replace(/\\s*-\\s*(Тур|Round)\\s*\\d+/gi, '')
                                   .replace(/Новое окно/gi, '')
                                   .replace(/\\s+/g, ' ')
                                   .trim();

                    // --- ДАТА ---
                    const dateEl = document.querySelector('.duelParticipant__startTime, .duelParticipant__startTime div');
                    const dateStr = dateEl ? dateEl.textContent.trim() : null;

                    // --- СЧЕТ (DOM + Title) ---
                    let ftHome = null, ftAway = null;
                    const scoreWrapper = document.querySelector('.detailScore__wrapper');
                    if (scoreWrapper) {
                        const spans = scoreWrapper.querySelectorAll('span');
                        if (spans.length >= 3) {
                            ftHome = spans[0].textContent.trim();
                            ftAway = spans[2].textContent.trim();
                        } else {
                            const parts = scoreWrapper.textContent.replace(/\\s+/g, '').split('-');
                            if (parts.length >= 2) { ftHome = parts[0]; ftAway = parts[1]; }
                        }
                    }

                    if (!ftHome || !ftAway) {
                        const hScore = document.querySelector('.detailScore__home');
                        const aScore = document.querySelector('.detailScore__away');
                        if (hScore && aScore) { ftHome = hScore.textContent.trim(); ftAway = aScore.textContent.trim(); }
                    }

                    if (!ftHome || !ftAway) {
                        const titleMatch = document.title.match(/(\\d+)\\s*[-:]\\s*(\\d+)/);
                        if (titleMatch) { ftHome = titleMatch[1]; ftAway = titleMatch[2]; }
                    }

                    return {
                        home: homeEl ? homeEl.textContent.trim() : null,
                        away: awayEl ? awayEl.textContent.trim() : null,
                        league: league,
                        dateStr: dateStr,
                        ftHome: ftHome,
                        ftAway: ftAway
                    };
                }''')

                if dom_data:
                    result['home_team'] = dom_data.get('home') or ""
                    result['away_team'] = dom_data.get('away') or ""
                    result['home_score'] = dom_data.get('ftHome') or ""
                    result['away_score'] = dom_data.get('ftAway') or ""
                    result['Liga'] = dom_data.get('league') or ""

                    if not result['match_date'] and dom_data.get('dateStr'):
                        try:
                            dt = datetime.strptime(dom_data['dateStr'], '%d.%m.%Y %H:%M')
                            result['match_date'] = dt.strftime('%Y.%m.%d %H:%M')
                        except: result['match_date'] = dom_data['dateStr']
            except Exception as e:
                logger.warning(f"   [{match_id}] DOM Error: {e}")
            finally:
                if page: await page.close()

            # 4. Статус
            if status_code:
                result['status_text'] = STATUS_MAP.get(status_code, 'Другой')
            elif result['home_score'] != "":
                result['status_text'] = 'Завершен'
            else:
                result['status_text'] = 'Неизвестно'

            logger.info(f"   [{match_id}] {result['Liga']} | {result['home_team']} {result['home_score']}-{result['away_score']} {result['away_team']} (HT: {result['ht_home_score']}-{result['ht_away_score']})")
            await asyncio.sleep(random.uniform(0.3, 0.6))
            return result

# ============================================================================
# MAIN
# ============================================================================

async def main():
    if not os.path.exists(INPUT_EXCEL):
        logger.error(f"❌ Файл не найден: {INPUT_EXCEL}")
        sys.exit(1)

    df = pd.read_excel(INPUT_EXCEL, sheet_name=SHEET_NAME)
    id_col = df.columns[0] if MATCH_ID_COLUMN is None else MATCH_ID_COLUMN

    target_columns = [
        'match_id', 'Liga', 'home_team', 'away_team',
        'home_score', 'away_score', 'ht_home_score',
        'ht_away_score', 'status_text', 'match_date'
    ]

    for col in target_columns:
        if col not in df.columns: df[col] = ""

    base, ext = os.path.splitext(INPUT_EXCEL)
    output_path = f"{base}{OUTPUT_SUFFIX}{ext}"

    resolver = MatchResolver()

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True, args=['--no-sandbox'])
        try:
            context = await resolver.bootstrap(browser)
            base_headers = {'X-Fsign': resolver.feed_sign, 'Accept': '*/*', 'User-Agent': USER_AGENT}
            sem = asyncio.Semaphore(MAX_CONCURRENT)

            tasks_meta = []
            for idx, row in df.iterrows():
                if LIMIT and len(tasks_meta) >= LIMIT: break

                # Resume-логика: ищем мусор в лиге ("Тур" или "окно") или пустой счет
                has_score = pd.notna(df.at[idx, 'home_score']) and str(df.at[idx, 'home_score']).strip() != ""
                liga_val = str(df.at[idx, 'Liga']).strip()
                has_dirty_league = "Тур" in liga_val or "окно" in liga_val.lower() or liga_val == ""

                if has_score and not has_dirty_league:
                    continue

                mid = str(row[id_col]).strip()
                if len(mid) > 3: tasks_meta.append((idx, mid))

            logger.info(f"🎯 К обработке: {len(tasks_meta)}")
            if not tasks_meta:
                save_with_widths(df, output_path, target_columns)
                return

            total, processed = len(tasks_meta), 0
            for start in range(0, total, CHUNK_SIZE):
                chunk = tasks_meta[start:start + CHUNK_SIZE]
                futures = [resolver.fetch_one(context, mid, base_headers, sem) for _, mid in chunk]
                results = await asyncio.gather(*futures, return_exceptions=True)

                for (idx, mid), res in zip(chunk, results):
                    processed += 1
                    if isinstance(res, dict):
                        for col in target_columns:
                            df.at[idx, col] = res.get(col, "")
                    else:
                        df.at[idx, 'status_text'] = 'Error'

                if processed % SAVE_EVERY == 0 or processed == total:
                    save_with_widths(df, output_path, target_columns)
                    logger.info(f"💾 Сохранено ({processed}/{total})")

            logger.info(f"🎉 ГОТОВО! Файл: {output_path}")

        finally:
            await browser.close()

if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        logger.warning("\n⚠️ Прервано")
    except Exception as e:
        logger.error(f"💥 Ошибка: {e}", exc_info=True)
