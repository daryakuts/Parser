from __future__ import annotations
import requests
from bs4 import BeautifulSoup
import os
import pandas
from io import BytesIO
from sqlalchemy import create_engine
import xlrd
import re
from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
import openpyxl


class html_converter:

    def __init__(self, data_start, data_end, link):
        self.data_start = data_start
        self.data_end = data_end
        self.link = link

    def get_page(self, number_page=None):
        """Загрузить страницу. Если number_page задан — с пагинацией."""
        parametr = None
        if number_page is not None:
            parametr = {'page': f'page-{number_page}'}
        response = self.open_link(self.link, parametr)
        if response is None:
            return None
        return BeautifulSoup(response.text, 'html.parser')

    def get_links_with_file(self, soup: BeautifulSoup):
        elements = soup.find_all('div', class_='accordeon-inner__item')
        results = []

        for item in elements:
            link_tag = None
            for a in item.find_all('a', href=True):
                href = a['href']
                # проверяем .xls/.xlsx где угодно в URL — не только в конце
                if '.xls' in href.lower() or '/oil_xls/' in href:
                    link_tag = a
                    break
            if link_tag is None:
                continue

            href = link_tag['href']
            date_tag = item.find('span')
            if date_tag is None:
                continue

            m = re.search(r'(\d{2})\.(\d{2})\.(\d{4})', date_tag.text)
            if not m:
                continue
            year = int(m.group(3))
            if not (self.data_start <= year <= self.data_end):
                continue

            url = href if href.startswith('http') else 'https://spimex.com' + href
            results.append({'date': year, 'link': url})

        return results

    def download_all_xlsx(self):
        links = self.get_all_links(max_pages=30)
        print(f'\n=== Всего найдено ссылок: {len(links)} ===')
        for i, item in enumerate(links):
            print(f'{i:3}  {item["date"]!r:40}  {item["link"]}')
        return links

    def get_all_links(self, start_page=19, max_pages=422, empty_streak_limit=5):
        all_links = []
        seen = set()
        empty_streak = 0

        for page in range(start_page, max_pages + 1):
            soup = self.get_page(page)
            if soup is None:
                print(f'Страница {page}: не загрузилась')
                empty_streak += 1
                if empty_streak >= empty_streak_limit:
                    print(f'{empty_streak} неудач подряд — стоп')
                    break
                continue

            links = self.get_links_with_file(soup)
            if not links:
                print(f'Страница {page}: XLS нет')
                empty_streak += 1
                if empty_streak >= empty_streak_limit:
                    print(f'{empty_streak} пустых подряд — стоп')
                    break
                continue

            empty_streak = 0

            new_links = [l for l in links if l['link'] not in seen]
            if new_links:
                for l in new_links:
                    seen.add(l['link'])
                all_links.extend(new_links)

            years = [l['date'] for l in links]
            print(f'Страница {page}: +{len(new_links)} '
                  f'(всего {len(all_links)}), годы {min(years)}..{max(years)}')

            if max(years) < self.data_start:
                print(f'Все годы на странице < {self.data_start} — стоп')
                break

        return all_links

    def open_link(self, link, parametr = None):

        try:
            response = requests.get(link, params=parametr, timeout=300)
        except requests.exceptions.RequestException as e:
            print(f'Ошибка сети при запросе {e}')
            return None
        if response.status_code != 200:
            print(f'Ошибка API (Код {response.status_code}). Ответ сервера:\n{response.text[:200]}')
            return None
        return response


if __name__=='__main__':
    test = html_converter(2023, 2026, 'https://spimex.com/markets/oil_products/trades/results/')
    links = test.get_all_links(start_page=19, max_pages=422, empty_streak_limit=5)
    print(f'\nВсего: {len(links)}')