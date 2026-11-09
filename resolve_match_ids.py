"""
🎯 FLASHSCORE SCORE RESOLVER v3.3 (API DE/DF Edition)

Собирает ТОЛЬКО:
match_id, Liga, home_team, away_team, home_score, away_score,
ht_home_score, ht_away_score, status_text, match_date

Ключевые изменения:
  ✅ ИТОГОВЫЙ СЧЕТ читается из API dc_1_ (ключи DE и DF) - 100% надежно
  ✅ Учет статуса матча (Не захватывает время как счет!)
  ✅ Исправлен синтаксис Python
  ✅ Защита от PermissionError при сохранении Excel
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

LIMIT = None
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
# ФУНКЦИЯ СОХРАНЕНИЯ EXCEL
# ============================================================================

def save_with_widths(df: pd.DataFrame, path: str, columns: list):
    try:
        df[columns].to_excel(path, index=False)
        wb = load_workbook(path)
        ws = wb.active

        # Pylance: wb.active может быть None — проверяем явно
        if ws is None:
            logger.warning(f"⚠️ В файле {path} нет активного листа, форматирование пропущено")
            wb.save(path)
            return

        for i, col in enumerate(columns, start=1):
            letter = get_column_letter(i)
            ws.column_dimensions[letter].width = COLUMN_WIDTHS.get(col, 15)

        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

        ws.freeze_panes = "A2"

        score_cols = {'home_score', 'away_score', 'ht_home_score', 'ht_away_score'}
        for i, col in enumerate(columns, start=1):
            if col in score_cols:
                letter = get_column_letter(i)
                for cell in ws[letter][1:]:
                    cell.alignment = Alignment(horizontal='center')

        wb.save(path)
    except PermissionError:
        logger.error(f"❌ Ошибка сохранения: закройте файл {path} в Excel!")
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
            # 1. Переходим на главную
            await page.goto(
                f"{BASE_DOMAIN}/football/",
                wait_until='domcontentloaded',
                timeout=60000,
            )

            # 2. Ждём, пока утихнут все фоновые запросы (редиректы)
            try:
                await page.wait_for_load_state('networkidle', timeout=10000)
            except Exception:
                # networkidle может не сработать из-за long-polling.
                # Не критично — идём дальше.
                pass

            # 3. Даём JS инициализироваться
            await asyncio.sleep(2)

            # 4. Retry-логика для evaluate (если контекст уничтожен)
            env = None
            for attempt in range(1, 4):
                try:
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
                    break  # успех — выходим из цикла
                except Exception as e:
                    if attempt == 3:
                        raise  # последняя попытка — пробрасываем
                    logger.debug(
                        f"⚠️ Bootstrap evaluate attempt {attempt}/3 failed: {e}"
                    )
                    # Ждём завершения навигации и пробуем снова
                    try:
                        await page.wait_for_load_state('networkidle', timeout=5000)
                    except Exception:
                        pass
                    await asyncio.sleep(2)

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

            # ========================================================================
            # 1. API dc_1_ (Дата, Статус и ИТОГОВЫЙ СЧЕТ из DE/DF)
            # ========================================================================
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
                            except Exception as e:
                                logger.debug(f"Date parse error for {match_id}: {e}")
                        if kv.get('DA'):
                            status_code = kv['DA']

                        # 🎯 ИТОГОВЫЙ СЧЕТ ИЗ API (ключи DE и DF)
                        # Читаем только если матч уже начался (LIVE или Завершен)
                        if status_code in ['2', '3', '6', '7', '31']:
                            if kv.get('DE') is not None and kv.get('DF') is not None:
                                result['home_score'] = kv['DE']
                                result['away_score'] = kv['DF']
            except Exception as e:
                logger.debug(f"API dc_1_ Error for {match_id}: {e}")

            # ========================================================================
            # 2. API df_sui_1_ (Счет первого тайма)
            # ========================================================================
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
            except Exception as e:
                logger.debug(f"API df_sui_1_ Error for {match_id}: {e}")

            # ========================================================================
            # 3. DOM (Только команды, лига, дата - БЕЗ счета!)
            # ========================================================================
            page = None
            try:
                page = await context.new_page()
                await page.goto(f"{self.base_domain}/match/{match_id}/", wait_until='domcontentloaded', timeout=20000)
                await asyncio.sleep(1.5)

                dom_data = await page.evaluate(r'''() => {
                    const homeEl = document.querySelector('.duelParticipant__home .participant__participantName');
                    const awayEl = document.querySelector('.duelParticipant__away .participant__participantName');

                    // --- ЛИГА ---
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

                    league = league.replace(/\s*-\s*(Тур|Round)\s*\d+/gi, '')
                                   .replace(/Новое окно/gi, '')
                                   .replace(/\s+/g, ' ')
                                   .trim();

                    // --- ДАТА ---
                    const dateEl = document.querySelector('.duelParticipant__startTime, .duelParticipant__startTime div');
                    const dateStr = dateEl ? dateEl.textContent.trim() : null;

                    return {
                        home: homeEl ? homeEl.textContent.trim() : null,
                        away: awayEl ? awayEl.textContent.trim() : null,
                        league: league,
                        dateStr: dateStr
                    };
                }''')

                if dom_data:
                    result['home_team'] = dom_data.get('home') or ""
                    result['away_team'] = dom_data.get('away') or ""
                    result['Liga'] = dom_data.get('league') or ""

                    # Фолбэк для даты из DOM
                    if not result['match_date'] and dom_data.get('dateStr'):
                        try:
                            dt = datetime.strptime(dom_data['dateStr'], '%d.%m.%Y %H:%M')
                            result['match_date'] = dt.strftime('%Y.%m.%d %H:%M')
                        except:
                            result['match_date'] = dom_data['dateStr']

            except Exception as e:
                logger.warning(f"   [{match_id}] DOM Error: {e}")
            finally:
                if page: await page.close()

            # ========================================================================
            # 4. Статус матча (с учетом наличия счета)
            # ========================================================================
            if status_code:
                result['status_text'] = STATUS_MAP.get(status_code, 'Другой')
            elif result['home_score'] != "":
                # Если счет есть, но статус не пришел из API - считаем завершенным
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
        if col not in df.columns:
            df[col] = ""

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
                if LIMIT and len(tasks_meta) >= LIMIT:
                    break

                has_score = pd.notna(df.at[idx, 'home_score']) and str(df.at[idx, 'home_score']).strip() != ""
                has_league = pd.notna(df.at[idx, 'Liga']) and str(df.at[idx, 'Liga']).strip() != ""

                if has_score and has_league:
                    continue

                mid = str(row[id_col]).strip()
                if len(mid) > 3:
                    tasks_meta.append((idx, mid))

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
